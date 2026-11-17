#!/usr/bin/env python3
"""puppa – vanliga M365-adminuppgifter i PowerShell med några knapptryck. Windows, macOS och Linux.

  puppa                    meny
  puppa kunder             sparade kunder (tenant + konto)
  puppa kund NAMN          växla kund
  puppa ny                 lägg till en kund (tenant, valfritt konto)
  puppa pnp [KUND]         sätt upp SharePoint/PnP för en kund (Entra ID-app eller befintligt ClientId)
  puppa test [KUND]        testa PnP-anslutningen
  puppa skal [KUND]        pwsh-skal som redan kört Connect-PnPOnline
  puppa recept             lista recept
  puppa kör RECEPT [Namn=värde …] [--ja]   kör ett recept (frågar efter det som saknas)
  puppa logga-ut           logga ut från PnP, Exchange, Teams och Graph (rensar sparade inloggningar)
  puppa hej

Inga kunder behöver skapas i förväg: Exchange-, Teams- och Graph-recept loggar in direkt och sparar
tenant + konto som en kund, så att du kan växla mellan kunder. SharePoint-recept går via PnP PowerShell,
som kräver en Entra ID-app per tenant – Puppa frågar efter den första gången.

Flaggor: -n/--torr visar kommandon utan att köra något · --no-anim stänger av animationer.
Data: ~/.puppa (kunder.json, inga hemligheter). Certifikatlösenord: SecretManagement-valvet "Puppa".
Egna recept: ~/.puppa/recipes/*.ps1. Enbart standardbiblioteket, Python 3.8+, kräver pwsh (PowerShell 7).
"""
from __future__ import annotations

import base64
import json
import os
import queue
import random
import re
import secrets
import shutil
import signal
import subprocess
import sys
import tempfile
import threading
import time
from datetime import datetime
from pathlib import Path

VERSION = "1.2.2"
HOME = Path.home()
DIR = Path(os.environ.get("PUPPA_DIR") or HOME / ".puppa")
os.environ["PUPPA_DIR"] = str(DIR)
CUSTOMERS = DIR / "kunder.json"
LEGACY_PROFILES = DIR / "profiles.json"
CONFIG = DIR / "config.json"
CERTS = DIR / "certs"
USER_RECIPES = DIR / "recipes"
APP_DIR = Path(os.path.realpath(__file__)).parent
RECIPES = APP_DIR / "recipes"
VAULT = "Puppa"
IS_WIN = os.name == "nt"
IS_MAC = sys.platform == "darwin"
if IS_WIN:
    _PF = os.environ.get("ProgramFiles", r"C:\Program Files")
    _LA = os.environ.get("LOCALAPPDATA", str(HOME / "AppData" / "Local"))
    EXTRA_PATH = [os.path.join(_PF, "PowerShell", "7"), os.path.join(_LA, "Microsoft", "WindowsApps")]
else:
    EXTRA_PATH = ["/opt/homebrew/bin", "/home/linuxbrew/.linuxbrew/bin", "/usr/local/bin",
                  "/usr/local/microsoft/powershell/7", "/snap/bin", "/usr/bin", "/bin"]
os.environ["PATH"] = os.pathsep.join(dict.fromkeys(os.environ.get("PATH", "").split(os.pathsep) + EXTRA_PATH))


def _win_console():
    """UTF-8 + ANSI-färger i Windows-konsolen (Windows Terminal, conhost, VS Code)."""
    for stream in (sys.stdout, sys.stderr, sys.stdin):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    try:
        import ctypes
        k = ctypes.windll.kernel32
        k.SetConsoleOutputCP(65001)
        k.SetConsoleCP(65001)
        h = k.GetStdHandle(-11)
        mode = ctypes.c_uint32()
        if k.GetConsoleMode(h, ctypes.byref(mode)):
            k.SetConsoleMode(h, mode.value | 0x0004)       # ENABLE_VIRTUAL_TERMINAL_PROCESSING
    except Exception:
        pass


if IS_WIN:
    _win_console()


def pwsh_exe():
    """PowerShell 7: pwsh, annars förhandsversionen (Homebrew-casken powershell@preview heter pwsh-preview)."""
    return shutil.which("pwsh") or shutil.which("pwsh-preview") or "pwsh"


def ps_args(script, *flags):
    """pwsh med -EncodedCommand: inga citat-/escape-problem på varken Windows eller Unix."""
    enc = base64.b64encode(script.encode("utf-16-le")).decode()
    return [pwsh_exe(), "-NoLogo", "-NoProfile", *flags, "-EncodedCommand", enc]


def detached():
    """Arbetsprocessen ska inte få Ctrl+C från terminalen – Puppa avbryter den själv."""
    return {"creationflags": 0x00000200} if IS_WIN else {"start_new_session": True}


def copy_text(s):
    for cmd in (["pbcopy"], ["clip"], ["wl-copy"], ["xclip", "-selection", "clipboard"]):
        exe = shutil.which(cmd[0])
        if exe:
            try:
                subprocess.run([exe] + cmd[1:], input=s, text=True, timeout=5)
                return True
            except Exception:
                return False
    return False


def open_file(path):
    code = shutil.which("code")
    try:
        if code:
            subprocess.Popen([code, str(path)])
        elif IS_WIN:
            os.startfile(str(path))                              # noqa – finns bara på Windows
        elif IS_MAC:
            subprocess.Popen(["open", "-t", str(path)])
        else:
            subprocess.Popen([shutil.which("xdg-open") or "xdg-open", str(path)])
    except Exception:
        out(dim(f"  Öppna själv: {path}"))


# ───────────────────────────── terminal ─────────────────────────────
TTY = sys.stdout.isatty()
ITTY = sys.stdin.isatty()
COLOR = TTY and "NO_COLOR" not in os.environ
ANIM = TTY and COLOR and not os.environ.get("PUPPA_NO_ANIM") and not os.environ.get("CI")
DRY = False
DIRTY = False      # något har skrivits ut som användaren bör hinna läsa innan skärmen töms


def c(code, s):
    return f"\033[{code}m{s}\033[0m" if COLOR else s


def green(s): return c("38;5;42", s)
def yellow(s): return c("38;5;214", s)
def red(s): return c("38;5;203", s)
def dim(s): return c("2", s)
def bold(s): return c("1", s)
def accent(s): return c("38;5;75", s)
def cyan(s): return c("38;5;80", s)


ANSI = re.compile(r"\x1b\[[0-9;?]*[a-zA-Z]")


def vlen(s):
    return len(ANSI.sub("", s))


def pad(s, n):
    return s + " " * max(0, n - vlen(s))


def cut(s, n):
    s = ANSI.sub("", s).replace("\t", " ").rstrip()
    return s if len(s) <= n else s[: max(0, n - 1)] + "…"


def width():
    return max(20, shutil.get_terminal_size((100, 30)).columns)


def height():
    return max(10, shutil.get_terminal_size((100, 30)).lines)


def fit(s, w=None):
    """Klipp en (färgad) rad så att den aldrig radbryts i terminalen."""
    w = max(8, (w or width()) - 1)
    if vlen(s) <= w:
        return s
    out, n, i = [], 0, 0
    while i < len(s) and n < w - 1:
        m = ANSI.match(s, i)
        if m:
            out.append(m.group(0))
            i = m.end()
            continue
        out.append(s[i])
        n += 1
        i += 1
    return "".join(out) + ("\033[0m" if COLOR else "") + "…"


def out(s=""):
    print(fit(s) if s else "")


def wrap(msg, w, maxl=4):
    lines, cur = [], ""
    for word in ANSI.sub("", msg).split():
        if cur and len(cur) + 1 + len(word) > w:
            lines.append(cur)
            cur = word
        else:
            cur = (cur + " " + word).strip()
    if cur:
        lines.append(cur)
    if len(lines) > maxl:
        lines = lines[:maxl]
        lines[-1] = cut(lines[-1] + " …", w)
    return [cut(l, w) for l in lines] or [""]


def hard_wrap(s, w):
    """Bryt en lång rad i bitar om w tecken (för kommandon som måste synas i sin helhet)."""
    s = s.replace("\t", "    ")
    if len(s) <= w:
        return [s]
    parts = [s[:w]]
    s = s[w:]
    while s:
        parts.append(s[: w - 2])
        s = s[w - 2:]
    return parts


class Live:
    """Ritar om ett block rader på plats, även efter att fönstret ändrat storlek."""

    def __init__(self):
        self.prev = []

    def _up(self):
        w = width()
        return sum(max(1, -(-n // w)) for n in self.prev)

    def draw(self, lines):
        w = width()
        lines = [fit(l, w) for l in lines]
        sys.stdout.write("\r" + "\033[F" * self._up() + "\033[J" + "".join(l + "\n" for l in lines))
        sys.stdout.flush()
        self.prev = [vlen(l) for l in lines]

    def clear(self):
        sys.stdout.write("\r" + "\033[F" * self._up() + "\033[J")
        sys.stdout.flush()
        self.prev = []


# ───────────────────────────── Puppa, maskoten ─────────────────────────────
# Puppa är ett tjockt P. Ansiktet bor i P:ets mage.
EYES = {"happy": " ◕ ◕ ", "blink": " ─ ─ ", "left": "◕ ◕  ", "right": "  ◕ ◕", "wait": " ◕ ◕ ",
        "sad": " ◕ ◕ ", "proud": " ^ ^ ", "wow": " ◕ ◕ ", "worry": " ◔ ◔ ", "sleep": " ─ ─ "}
MOUTH = {"happy": "  ‿  ", "blink": "  ‿  ", "left": " ‿   ", "right": "   ‿ ", "wait": "  ─  ",
         "sad": "  ⁀  ", "proud": "  ◡  ", "wow": "  o  ", "worry": "  ~  ", "sleep": "  ᶻ  "}
ART = ["████████▄", "██{e}██", "██{m}██", "████████▀", "██       ", "▀▀       "]
SHADE = ["38;5;117", "38;5;75", "38;5;69", "38;5;63"]
ART_W = 9
LOGO = "P›"
JOKES = [
    "Jag heter Puppa. Det står för PowerShell-uppgifter på autopilot. Typ.",
    "Get-PnPTenantSite: för att du aldrig vet hur många sajter som finns förrän du frågar.",
    "-WhatIf är mitt mellannamn. -Confirm är mitt efternamn.",
    "Jag har inga ben, bara en stapel. Men jag står stadigt.",
    "Det finns två sorters admins: de som tar backup, och de som kommer att göra det.",
    "Connect-ExchangeOnline: för att en inloggning om dagen aldrig räcker.",
    "Mitt favoritrecept? Exportera till Excel. Chefen älskar Excel.",
    "AzureAD-modulen? Vi pratar inte om AzureAD-modulen.",
]


def puppa_art(mood="happy"):
    e, m = EYES.get(mood, EYES["happy"]), MOUTH.get(mood, MOUTH["happy"])
    sh = SHADE if mood != "sad" else ["38;5;110", "38;5;67", "38;5;61", "38;5;60"]

    def mid(k, inner):
        return c(sh[k], "██") + bold(inner) + c(sh[k], "██")
    return [c(sh[0], ART[0]), mid(1, e), mid(2, m), c(sh[2], ART[3]), c(sh[3], ART[4]), c(sh[3], ART[5])]


def frame(msg="", mood="happy", lift=0, top="", sub=""):
    """Maskot + pratbubbla. Under 44 kolumner döljs maskoten."""
    w = width()
    if w < 44:
        lines = []
        if msg:
            for k, l in enumerate(wrap(msg, w - 6, 4)):
                lines.append((f"  {accent(bold(LOGO))} " if k == 0 else "     ") + l)
        if sub:
            lines.append("     " + dim(cut(sub, w - 7)))
        return lines
    rows = [""] * 7
    for k, r in enumerate(puppa_art(mood)):
        rows[k + 1 - lift] = r
    if top and not rows[0]:
        rows[0] = top
    bub = []
    if msg:
        lines = wrap(msg, max(14, min(60, w - 20)), 4)
        iw = max(len(l) for l in lines)
        bub = [dim("╭" + "─" * (iw + 2) + "╮")] + \
              [dim("│ ") + l + " " * (iw - len(l)) + dim(" │") for l in lines] + \
              [dim("╰" + "─" * (iw + 2) + "╯")]
    res = []
    for k in range(7):
        line = "  " + pad(rows[k], ART_W + 1)
        if 0 <= k - 1 < len(bub):
            line += " " + (accent("◂ ") if k == 2 else "  ") + bub[k - 1]
        elif sub and k - 1 == len(bub):
            line += "    " + dim(cut(sub, w - ART_W - 9))
        res.append(fit(line, w))
    while res and not ANSI.sub("", res[-1]).strip():
        res.pop()
    while res and not ANSI.sub("", res[0]).strip():
        res.pop(0)
    return res


def say(msg, mood="happy", sub=""):
    global DIRTY
    DIRTY = True
    if TTY:
        print("\n".join(frame(msg, mood, sub=sub)))
    else:
        print(f"  {LOGO} {msg}")


def idle_mood(t):
    """Blinkar då och då, tittar runt ibland."""
    if t % 4.1 < 0.16:
        return "blink"
    phase = int(t / 2.3) % 7
    return {2: "left", 5: "right"}.get(phase, "happy")


def celebrate(msg):
    global DIRTY
    DIRTY = True
    if not ANIM:
        return say(msg, "proud")
    cols = ["38;5;117", "38;5;214", "38;5;80", "38;5;212", "38;5;227"]
    live = Live()
    sys.stdout.write("\033[?25l")
    try:
        for f in range(14):
            top = "".join(c(random.choice(cols), random.choice("✦✧⋆·*")) if random.random() < .35 else " "
                          for _ in range(ART_W + 1))
            live.draw(frame(msg, "proud" if f % 5 else "blink", lift=f % 2, top=top))
            time.sleep(0.09)
        live.draw(frame(msg, "proud"))
    finally:
        sys.stdout.write("\033[?25h")
        sys.stdout.flush()


# ───────────────────────────── input ─────────────────────────────
class Cancel(Exception):
    pass


class Keys:
    """Rå tangentläsning. Windows: msvcrt. macOS/Linux: termios + select. read() ger None vid timeout,
    så menyerna kan animera och rita om när fönstret ändrat storlek (även där SIGWINCH saknas)."""

    def __enter__(self):
        if IS_WIN:
            import msvcrt
            self.m = msvcrt
        else:
            import termios
            import tty
            self.fd = sys.stdin.fileno()
            self.old = termios.tcgetattr(self.fd)
            tty.setcbreak(self.fd)
        sys.stdout.write("\033[?25l")
        sys.stdout.flush()
        return self

    def __exit__(self, *a):
        if not IS_WIN:
            import termios
            termios.tcsetattr(self.fd, termios.TCSADRAIN, self.old)
        sys.stdout.write("\033[?25h")
        sys.stdout.flush()

    def read(self, timeout=0.12):
        if IS_WIN:
            end = time.time() + timeout
            while time.time() < end:
                if self.m.kbhit():
                    ch = self.m.getwch()
                    if ch in ("\x00", "\xe0"):
                        return {"H": "\x1b[A", "P": "\x1b[B", "K": "\x1b[D", "M": "\x1b[C"}.get(self.m.getwch(), "")
                    return ch
                time.sleep(0.02)
            return None
        import select
        r, _, _ = select.select([self.fd], [], [], timeout)
        return os.read(self.fd, 8).decode(errors="ignore") if r else None


def new_screen():
    """Töm skärmen inför en meny. Har något just skrivits ut får man läsa klart först."""
    global DIRTY
    if DIRTY:
        pause()
    DIRTY = False
    if TTY:
        sys.stdout.write("\033[H\033[2J")
        sys.stdout.flush()


def fuzzy_score(query, text, loose=True):
    """Fuzzy-match: alla ord i frågan måste finnas i texten, som delsträng (bäst) eller som bokstäver i ordning.
    Högre = bättre. None = ingen träff."""
    text = ANSI.sub("", text).lower()
    total = 0
    for word in query.lower().split():
        i = text.find(word)
        if i >= 0:
            start = i == 0 or not text[i - 1].isalnum()
            total += 100 + (40 if start else 0) - min(i, 30)
            continue
        if len(word) < 3 or not loose:       # korta sökord / långa texter: bara som delsträng
            return None
        pos, gaps, last = 0, 0, -1
        for ch in word:
            j = text.find(ch, pos)
            if j < 0:
                return None
            if last >= 0:
                gaps += j - last - 1
            last, pos = j, j + 1
        if gaps > 2 * len(word):             # för utspritt för att vara det man menade
            return None
        total += max(1, 60 - gaps * 3)
    return total


def fuzzy_rank(query, items, key):
    if not query.strip():
        return list(items)
    scored = [(fuzzy_score(query, key(x)), n, x) for n, x in enumerate(items)]
    return [x for s, n, x in sorted([y for y in scored if y[0] is not None], key=lambda y: (-y[0], y[1]))]


def choose(rows, title="", msg="", mood="happy", sub="", status=None, start=None, footer=None, screen=False,
           search_rows=None):
    """Piltangent-meny. rows: ("head", text) eller dict(label, hint, value, badge, key).
    Siffra väljer direkt, q/esc avbryter (→ None). Börja skriva för att fuzzy-söka (search_rows = sök i en
    större lista, t.ex. alla recept). Maskoten blinkar medan den väntar.
    screen=True: menyn ritas på en tom skärm (navigeringsmenyer), annars där markören står (frågor i ett flöde)."""
    if not ITTY:
        raise Cancel()
    if screen:
        new_screen()
    hot = {r["key"] for r in rows if not isinstance(r, tuple) and r.get("key")}
    st = {"q": "", "pos": 0, "top": 0, "mood": None}

    def view():
        if st["q"]:
            pool = [r for r in (search_rows if search_rows is not None else rows) if not isinstance(r, tuple)]
            pool = [r for r in pool if r.get("value") is not None]
            q = st["q"]
            scored = []
            for n, r in enumerate(pool):
                a = fuzzy_score(q, r["label"])
                b = fuzzy_score(q, r["label"] + " " + ANSI.sub("", r.get("hint") or "") + " " + (r.get("search") or ""),
                                loose=False)
                if a is None and r.get("search"):
                    a = fuzzy_score(q, r["search"])
                sc = max(a * 2 if a is not None else -1, b if b is not None else -1)
                if sc >= 0:
                    scored.append((-sc, n, r))
            return [r for _, _, r in sorted(scored, key=lambda x: (x[0], x[1]))]
        return rows

    def numbering(vr):
        sel = [k for k, r in enumerate(vr) if not isinstance(r, tuple)]
        keys, n = {}, 0
        for k in sel:
            r = vr[k]
            if r.get("key") and not st["q"]:
                keys[r["key"]] = k
            else:
                n += 1
                if n <= 9:
                    keys[str(n)] = k
        return sel, keys

    vr = view()
    sel, keys = numbering(vr)
    if not sel:
        return None
    st["pos"] = sel[0]
    if start is not None:
        for k in sel:
            if vr[k]["value"] == start:
                st["pos"] = k
    live = Live()
    winch = {"flag": False}
    t0 = time.time()

    def draw(mood_now):
        w = width()
        lines = []
        if msg and not st["q"]:
            lines += frame(msg, mood_now, sub=sub)
        if status:
            lines += [""] + status()
        lines.append("")
        if st["q"]:
            lines.append("  " + accent("⌕ ") + bold(st["q"]) + accent("▏") + dim(f"  {len(sel)} träffar"))
        elif title:
            lines.append("  " + accent("▍") + bold(title))
        narrow = w < 58
        if not sel:
            lines.append("     " + dim("Inga träffar. Backsteg ändrar, esc rensar."))
        labels = [vr[k]["label"] for k in sel] or [""]
        lw = min(max(len(x) for x in labels) + 2, max(10, w - 14 if narrow else min(34, w // 2)))
        room = max(3, height() - len(lines) - 3)
        pos = st["pos"]
        top = min(max(st["top"], pos - room + 1), pos)
        top = max(0, min(top, max(0, len(vr) - room)))
        st["top"] = top
        if top > 0:
            lines.append(dim("     ↑ fler"))
        num = {k: n for n, k in keys.items()}
        anybadge = any(vr[j].get("badge") for j in sel)
        for k in range(top, min(len(vr), top + room)):
            r = vr[k]
            if isinstance(r, tuple):
                lines.append("   " + dim(cut(r[1].upper(), w - 5)))
                continue
            on = k == pos
            key = num.get(k, " ")
            label = cut(r["label"], lw - 1)
            label = bold(accent(label)) if on else label
            badge = r.get("badge") or ""
            if narrow:
                line = f"  {accent('❯') if on else ' '}{dim(key)} " + (pad(label, lw) + badge if anybadge else label)
            else:
                line = f"   {accent('❯') if on else ' '} {dim(key)}  " + pad(label, lw) + (pad(badge, 8) if anybadge else "")
                if r.get("hint") and w >= 64:
                    line += dim(ANSI.sub("", r["hint"]))
            lines.append(line)
        if top + room < len(vr):
            lines.append(dim("     ↓ fler"))
        lines.append("")
        if st["q"]:
            help_ = "↑/↓ · enter väljer · backsteg · esc rensar" if w >= 50 else "↑↓ enter · esc"
        else:
            help_ = ((footer or "↑/↓ · enter · siffra väljer · q tillbaka") + " · skriv för att söka") if w >= 72 else \
                (footer or "↑/↓ · enter · q tillbaka") if w >= 50 else "↑↓ enter · q"
        lines.append(dim("   " + help_))
        live.draw(lines)

    def refilter():
        nonlocal vr, sel, keys
        vr = view()
        sel, keys = numbering(vr)
        st["pos"] = sel[0] if sel else 0
        st["top"] = 0

    def on_winch(*_):
        winch["flag"] = True

    old_winch = signal.signal(signal.SIGWINCH, on_winch) if hasattr(signal, "SIGWINCH") else None
    size = shutil.get_terminal_size((100, 30))
    choice = None
    picked = None
    kb = Keys()
    try:
        kb.__enter__()
        draw(mood)
        while True:
            ch = kb.read(0.12)
            if ch is None:
                now = shutil.get_terminal_size((100, 30))
                if now != size:              # fönstret ändrat storlek (Windows har ingen SIGWINCH)
                    size = now
                    winch["flag"] = True
                m = mood
                if ANIM and not st["q"]:
                    im = idle_mood(time.time() - t0)
                    m = im if mood == "happy" else ("blink" if im == "blink" else mood)
                if winch["flag"] or m != st["mood"]:
                    winch["flag"] = False
                    st["mood"] = m
                    draw(m)
                continue
            idx = sel.index(st["pos"]) if st["pos"] in sel else 0
            if ch in ("\x1b[A", "\x1bOA"):
                if sel:
                    st["pos"] = sel[(idx - 1) % len(sel)]
            elif ch in ("\x1b[B", "\x1bOB"):
                if sel:
                    st["pos"] = sel[(idx + 1) % len(sel)]
            elif ch in ("\r", "\n"):
                if sel:
                    picked = vr[st["pos"]]
                    choice = picked["value"]
                    break
            elif ch in ("\x7f", "\x08"):
                if st["q"]:
                    st["q"] = st["q"][:-1]
                    refilter()
            elif ch == "\x1b":
                if st["q"]:
                    st["q"] = ""
                    refilter()
                else:
                    break
            elif ch in ("\x03", "\x04"):
                break
            elif not st["q"] and ch.isdigit() and ch in keys:
                picked = vr[keys[ch]]
                choice = picked["value"]
                break
            elif not st["q"] and ch in hot:
                picked = vr[keys[ch]]
                choice = picked["value"]
                break
            elif not st["q"] and ch == "q":
                break
            elif ch and not ch.startswith("\x1b") and ch.isprintable() and (st["q"] or not ch.startswith(" ")):
                st["q"] += ch                    # även inklistrat / snabbt skrivet (flera tecken på en gång)
                refilter()
            draw(st["mood"] or mood)
    except KeyboardInterrupt:
        choice = None
    finally:
        if old_winch is not None:
            signal.signal(signal.SIGWINCH, old_winch)
        kb.__exit__()
        live.clear()
        sys.stdout.flush()
    if choice is not None and not screen:
        out(f"   {accent('❯')} {bold(picked['label'])}")
    return choice


def ask_text(question, default="", required=False, check=None, secret=False):
    """Fråga efter en text. Enter = standardvärdet. Ctrl+C avbryter flödet."""
    if not ITTY:
        if required and not default:
            raise Cancel()
        return default
    w = width()
    for k, l in enumerate(wrap(question, w - 6, 6)):
        out(("  " + accent("?") + " " if k == 0 else "    ") + bold(l))
    if default:
        out("    " + dim("standard: " + cut(default, w - 16) + "  (enter behåller)"))
    while True:
        try:
            if secret:
                import getpass
                v = getpass.getpass("    › ")
            else:
                v = input("    " + accent("›") + " ")
        except (EOFError, KeyboardInterrupt):
            print()
            raise Cancel()
        v = v.strip()
        if not v:
            v = default
        if not v and required:
            out("    " + yellow("Det här behövs. Ctrl+C avbryter."))
            continue
        if check and v:
            err = check(v)
            if err:
                out("    " + yellow(err))
                continue
        return v


def confirm(question, default=False):
    if not ITTY:
        return default
    opts = "[J/n]" if default else "[j/N]"
    try:
        a = input(fit(f"  {accent('?')} {bold(question)} {dim(opts)}") + " ").strip().lower()
    except (EOFError, KeyboardInterrupt):
        print()
        return False
    if not a:
        return default
    return a in ("j", "ja", "y", "yes")


def pause():
    """Vänta på enter (eller esc/q) innan skärmen töms. Läser tangenter rått så att piltangenter inte syns som ^[[A."""
    if ITTY:
        print()
        sys.stdout.write(dim("   enter = tillbaka till menyn "))
        sys.stdout.flush()
        kb = Keys()
        try:
            kb.__enter__()
            while True:
                ch = kb.read(0.5)
                if ch is None:
                    continue
                if ch in ("\r", "\n", "\x1b", "q", "\x03", "\x04", " "):
                    break
        except KeyboardInterrupt:
            pass
        finally:
            kb.__exit__()
            print()


# ───────────────────────────── lagring ─────────────────────────────
GUID = re.compile(r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$")


def private(path):
    try:
        os.chmod(path, 0o700 if Path(path).is_dir() else 0o600)
    except OSError:
        pass


def load_json(path, default):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def save_json(path, data):
    DIR.mkdir(parents=True, exist_ok=True)
    private(DIR)
    tmp = Path(str(path) + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=1, ensure_ascii=False)
    private(tmp)
    os.replace(tmp, path)


def config():
    return load_json(CONFIG, {})


def set_config(**kw):
    cfg = config()
    cfg.update(kw)
    save_json(CONFIG, cfg)


# En "tenant" = en kunds tenant som du loggat in i: domän, tenant-id, konto och (valfritt) PnP-app.
def tenants():
    d = load_json(CUSTOMERS, None)
    if d is None:
        old = load_json(LEGACY_PROFILES, {}).get("profiles")
        if old:
            save_json(CUSTOMERS, {"tenants": old})
            return old
        return []
    return d.get("tenants", [])


def save_tenants(ts):
    save_json(CUSTOMERS, {"tenants": ts})


def find_tenant(name):
    if not name:
        return None
    n = name.lower()
    ts = tenants()
    for t in ts:
        if n in (t["name"].lower(), t.get("tenant", "").lower(), t.get("tenant_id", "").lower()):
            return t
    hits = [t for t in ts if t["name"].lower().startswith(n) or t.get("tenant", "").lower().startswith(n)]
    return hits[0] if len(hits) == 1 else None


def active_tenant():
    return find_tenant(config().get("active"))


def update_tenant(name, **kw):
    ts = tenants()
    for t in ts:
        if t["name"] == name:
            t.update(kw)
    save_tenants(ts)


def slug(s):
    return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-") or "tenant"


def unique_name(base, ts):
    names = {t["name"].lower() for t in ts}
    name, k = base, 2
    while name.lower() in names:
        name, k = f"{base} {k}", k + 1
    return name


def tenant_info(t):
    """'contoso' / 'contoso.onmicrosoft.com' / 'contoso.se' / en SharePoint-URL -> (tenant, sharepoint-prefix)."""
    t = t.strip().lower()
    t = re.sub(r"^https?://", "", t).split("/")[0]
    if t.endswith(".sharepoint.com"):
        pre = t.split(".")[0].replace("-admin", "")
        return pre + ".onmicrosoft.com", pre
    if GUID.match(t):
        return t, ""
    if "." not in t:
        return t + ".onmicrosoft.com", t
    if t.endswith(".onmicrosoft.com"):
        return t, t.split(".")[0]
    return t, ""


def sp_fields(prefix):
    return {"prefix": prefix, "admin_url": f"https://{prefix}-admin.sharepoint.com",
            "root_url": f"https://{prefix}.sharepoint.com"} if prefix else {}


def remember_login(who):
    """Efter en lyckad inloggning: spara/uppdatera tenanten (domän, id, konto) och gör den aktiv."""
    tid, upn, domain, org = (list(who) + ["", "", "", ""])[:4]
    if not domain and "@" in upn:
        domain = upn.split("@", 1)[1]
    ts = tenants()
    t = next((x for x in ts if tid and x.get("tenant_id", "").lower() == tid.lower()), None) or \
        next((x for x in ts if domain and x.get("tenant", "").lower() == domain.lower()), None)
    new = t is None
    now = datetime.now().isoformat(timespec="seconds")
    if new:
        base = org or (domain.split(".")[0].capitalize() if domain else "Tenant")
        t = dict(name=unique_name(base, ts), tenant=domain or tid, tenant_id=tid, upn=upn, client_id="",
                 auth="interactive", created=now)
        ts.append(t)
    else:
        if tid:
            t["tenant_id"] = tid
        if upn:
            t["upn"] = upn
        cur = t.get("tenant", "")
        if domain and (not cur or GUID.match(cur) or
                       (domain.endswith(".onmicrosoft.com") and not cur.endswith(".onmicrosoft.com"))):
            t["tenant"] = domain
    if not t.get("prefix") and t.get("tenant", "").endswith(".onmicrosoft.com"):
        t.update(sp_fields(t["tenant"].split(".")[0]))
    t["last"] = now
    save_tenants(ts)
    set_config(active=t["name"])
    return t, new


# Hemligheter (certifikatlösenord) i PowerShells SecretManagement-valv "Puppa" (SecretStore) – samma på alla OS.
SEC_PS = r"""
$ErrorActionPreference = 'Stop'
Import-Module Microsoft.PowerShell.SecretManagement, Microsoft.PowerShell.SecretStore
$V = $env:PUPPA_VAULT
function Open-Store { if ($env:PUPPA_STORE_PW) { Unlock-SecretStore -Password (ConvertTo-SecureString $env:PUPPA_STORE_PW -AsPlainText -Force) } }
switch ($env:PUPPA_OP) {
  'state' {
    $reg = [bool](Get-SecretVault -Name $V -ErrorAction SilentlyContinue)
    try { Open-Store; $a = (Get-SecretStoreConfiguration -ErrorAction Stop).Authentication; "ok|$a|$reg" }
    catch { if ("$_" -match 'Unlock-SecretStore|valid password|password is not') { "locked||$reg" } else { "none||$reg" } }
  }
  'setup' {
    if ($env:PUPPA_STORE_PW) {
      Set-SecretStoreConfiguration -Scope CurrentUser -Authentication Password -Interaction None -Confirm:$false `
        -Password (ConvertTo-SecureString $env:PUPPA_STORE_PW -AsPlainText -Force)
    } else { Set-SecretStoreConfiguration -Scope CurrentUser -Authentication None -Interaction None -Confirm:$false }
    if (-not (Get-SecretVault -Name $V -ErrorAction SilentlyContinue)) { Register-SecretVault -Name $V -ModuleName Microsoft.PowerShell.SecretStore }
    'ok'
  }
  'register' { Open-Store; if (-not (Get-SecretVault -Name $V -ErrorAction SilentlyContinue)) { Register-SecretVault -Name $V -ModuleName Microsoft.PowerShell.SecretStore }; 'ok' }
  'get' { Open-Store; Get-Secret -Name $env:PUPPA_NAME -Vault $V -AsPlainText }
  'set' { Open-Store; Set-Secret -Name $env:PUPPA_NAME -Vault $V -Secret $env:PUPPA_VAL; 'ok' }
  'del' { Open-Store; Remove-Secret -Name $env:PUPPA_NAME -Vault $V -ErrorAction SilentlyContinue; 'ok' }
}
"""
STORE_PW = None


def sec_call(op, name="", value=""):
    env = dict(os.environ, PUPPA_OP=op, PUPPA_VAULT=VAULT, PUPPA_NAME=name, PUPPA_VAL=value)
    if STORE_PW:
        env["PUPPA_STORE_PW"] = STORE_PW
    try:
        p = subprocess.run(ps_args(SEC_PS, "-NonInteractive"), capture_output=True, text=True, encoding="utf-8",
                           errors="replace", env=env, timeout=90, stdin=subprocess.DEVNULL)
    except (OSError, subprocess.TimeoutExpired) as e:
        return False, "", str(e)
    return p.returncode == 0, p.stdout.strip(), p.stderr.strip()


def store_state():
    ok, o, _ = sec_call("state")
    line = (o.splitlines() or ["none||False"])[-1]
    st, _, rest = line.partition("|")
    auth, _, reg = rest.partition("|")
    return st, auth, reg == "True"


def store_ready():
    """Valvet 'Puppa' finns och är upplåst. Skapar det (med eller utan lösenord) första gången."""
    global STORE_PW
    st, _, reg = store_state()
    if st == "locked":
        for _ in range(3):
            STORE_PW = ask_text("Lösenord till ditt SecretStore-valv", secret=True, required=True)
            st, _, reg = store_state()
            if st == "ok":
                break
            out("    " + yellow("Fel lösenord."))
        else:
            STORE_PW = None
            return False
    elif st == "none":
        say("Certifikatlösenordet sparas krypterat i PowerShells SecretStore, i ett valv som heter Puppa. "
            "Du har inget valv än, så jag skapar ett.", "happy")
        how = choose([dict(label="Utan eget lösenord", hint="enklast – krypterat för ditt användarkonto", value="none"),
                      dict(label="Med lösenord", hint="säkrare – Puppa frågar en gång per körning", value="pw")],
                     title="Hur ska valvet låsas?")
        if how is None:
            raise Cancel()
        if how == "pw":
            a = ask_text("Välj ett lösenord för valvet", secret=True, required=True)
            if ask_text("Samma lösenord igen", secret=True, required=True) != a:
                say("Lösenorden var olika. Försök igen.", "sad")
                return False
            STORE_PW = a
        ok, _, err = sec_call("setup")
        if not ok:
            say("Kunde inte skapa valvet: " + (err.splitlines() or ["okänt fel"])[-1], "sad")
            return False
        return True
    if not reg:
        ok, _, err = sec_call("register")
        if not ok:
            say("Kunde inte registrera valvet Puppa: " + (err.splitlines() or ["okänt fel"])[-1], "sad")
            return False
    return True


def secret_name(t):
    return "puppa-" + slug(t["name"])


def sec_set(t, value):
    if not store_ready():
        return False
    ok, _, _ = sec_call("set", secret_name(t), value)
    return ok


def sec_get(t):
    ok, o, _ = sec_call("get", secret_name(t))
    if not (ok and o):
        try:
            if not store_ready():
                return None
        except Cancel:
            return None
        ok, o, _ = sec_call("get", secret_name(t))
    return o.splitlines()[-1] if ok and o else None


def sec_del(t):
    sec_call("del", secret_name(t))


# ───────────────────────────── PowerShell ─────────────────────────────
def psq(v):
    """PowerShell-sträng med enkla citattecken."""
    return "'" + str(v).replace("'", "''") + "'"


SECRET_EXPR = "(ConvertTo-SecureString $env:PUPPA_SECRET -AsPlainText -Force)"
R = "\x1e"
WORKER_PS = r"""
$ErrorActionPreference = 'Continue'
$ProgressPreference = 'SilentlyContinue'
$PSStyle.OutputRendering = 'PlainText'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
[Console]::InputEncoding = [Text.Encoding]::UTF8
$global:PuppaLast = @()
$R = [char]0x1e
function global:Out-Puppa([string]$kind, [string]$text) { [Console]::Out.WriteLine("$R$kind " + ($text -replace "`r?`n", ' ')); [Console]::Out.Flush() }
Out-Puppa 'D' '1 0'
while ($true) {
  $buf = [System.Text.StringBuilder]::new()
  while ($true) {
    $l = [Console]::In.ReadLine()
    if ($null -eq $l) { exit 0 }
    if ($l -eq "${R}END") { break }
    [void]$buf.AppendLine($l)
  }
  $text = $buf.ToString()
  $nl = $text.IndexOf("`n")
  $w = [int]$text.Substring(0, $nl).Trim()
  $code = $text.Substring($nl + 1)
  $st = @{ ok = $true }
  $data = [System.Collections.Generic.List[object]]::new()
  try {
    $sb = [scriptblock]::Create($code)
    & $sb 2>&1 3>&1 | ForEach-Object {
      if ($_ -is [System.Management.Automation.ErrorRecord]) { $st.ok = $false; Out-Puppa 'E' $_.ToString() }
      elseif ($_ -is [System.Management.Automation.WarningRecord]) { Out-Puppa 'W' $_.Message }
      elseif ($null -ne $_) { $data.Add($_) }
    }
  } catch { $st.ok = $false; Out-Puppa 'E' $_.ToString() }
  finally { Remove-Item Env:PUPPA_SECRET -ErrorAction SilentlyContinue }
  if ($data.Count) {
    $global:PuppaLast = $data.ToArray()
    try { $data | Export-Clixml -LiteralPath (Join-Path $env:PUPPA_DIR 'senaste.xml') -Depth 2 } catch { }
    try { $data | Out-String -Stream -Width $w | ForEach-Object { [Console]::Out.WriteLine("${R}O " + $_) } }
    catch { Out-Puppa 'E' $_.ToString() }
  }
  Out-Puppa 'D' ("{0} {1}" -f [int]$st.ok, $data.Count)
}
"""


class Worker:
    """En långlivad pwsh-process: anslutningarna lever kvar mellan recepten.
    kinds = vilka anslutningar processen äger; conn delas mellan processerna (se Workers)."""

    def __init__(self, kinds=(), conn=None):
        self.p = None
        self.q = queue.Queue()
        self.kinds = set(kinds)
        self.conn = {} if conn is None else conn
        self.lock = threading.Lock()
        self.ready = threading.Event()

    def start(self):
        with self.lock:
            if self.p and self.p.poll() is None:
                return
            self.forget()
            self.ready.clear()
            DIR.mkdir(parents=True, exist_ok=True)
            private(DIR)
            self.q = queue.Queue()
            self.p = subprocess.Popen(ps_args(WORKER_PS, "-NonInteractive"),
                                      stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                      text=True, encoding="utf-8", errors="replace", bufsize=1, **detached())
            q, p = self.q, self.p

            def reader():
                for line in p.stdout:
                    q.put(line.rstrip("\n"))
                q.put(None)
            threading.Thread(target=reader, daemon=True).start()
        # första D-raden = redo
        while True:
            ln = self.q.get()
            if ln is None or ln.startswith(R + "D"):
                break
        self.ready.set()

    def forget(self):
        for k in self.kinds:
            self.conn.pop(k, None)

    def warm(self):
        threading.Thread(target=self.start, daemon=True).start()

    def stop(self):
        if self.p and self.p.poll() is None:
            try:
                self.p.stdin.close()
                self.p.wait(3)
            except Exception:
                self.p.kill()
        self.p = None
        self.forget()

    def restart(self):
        self.stop()
        self.start()

    def run(self, code, on_event, secret=None):
        """Kör kod; on_event(kind, text) får H(ost)/W/E/O-rader. -> (ok, antal objekt)."""
        if not self.p or self.p.poll() is not None:
            self.start()
        self.ready.wait(120)
        pre = ""
        if secret:
            pre = f"$env:PUPPA_SECRET = {psq(secret)}\n"
        payload = f"{max(30, width() - 6)}\n{pre}{code}\n{R}END\n"
        try:
            self.p.stdin.write(payload)
            self.p.stdin.flush()
        except (BrokenPipeError, OSError):
            on_event("E", "PowerShell-processen har stängts. Försök igen.")
            self.p = None
            return False, 0
        while True:
            ln = self.q.get()
            if ln is None:
                on_event("E", "PowerShell-processen avslutades oväntat.")
                self.p = None
                return False, 0
            if ln.startswith(R):
                kind, _, text = ln[1:].partition(" ")
                if kind == "D":
                    ok, _, n = text.partition(" ")
                    return ok == "1", int(n or 0)
                on_event(kind, text)
            elif ln.startswith(("#< CLIXML", "<Objs ")):
                continue                 # förloppsdata som pwsh skriver till stderr (CLIXML) – inget att visa
            else:
                on_event("H", ln)


class Workers:
    """PnP.PowerShell och ExchangeOnlineManagement/Graph/Teams laddar olika versioner av MSAL
    (Microsoft.Identity.Client). I samma process vinner den som laddas först och den andra kraschar
    ("Method not found … WithBroker"). Därför: PnP i en egen pwsh, allt annat i en annan."""
    PNP = re.compile(r"\b[A-Za-z]+-PnP[A-Za-z]")

    def __init__(self):
        self.conn = {}
        self.main = Worker(("exo", "teams", "graph"), self.conn)
        self.pnp = Worker(("pnp",), self.conn)

    def pick(self, code):
        return self.pnp if self.PNP.search(code) else self.main

    def run(self, code, on_event, secret=None):
        return self.pick(code).run(code, on_event, secret)

    def start(self):
        self.main.start()

    def warm(self):
        self.main.warm()          # PnP-processen startas först när något PnP-kommando körs

    def stop(self):
        self.main.stop()
        self.pnp.stop()

    def restart(self):
        self.stop()
        self.start()


WORKER = Workers()


class Runner:
    """Visar maskoten som väntar (blinkar, tittar runt) och de senaste raderna medan något körs."""
    DEVICE = re.compile(r"(https?://\S*(?:devicelogin|deviceauth)\S*).*?\b(?:code|kod)\b\s+([A-Z0-9-]{6,12})", re.I)

    def __init__(self, title):
        self.title = title
        self.lines = []      # (kind, text)
        self.tail = []
        self.msg = None
        self.stop = threading.Event()
        self.lock = threading.Lock()
        self.live = Live()
        self.t0 = time.time()
        self.th = threading.Thread(target=self.loop, daemon=True)

    def event(self, kind, text):
        with self.lock:
            self.lines.append((kind, text))
            if kind in ("H", "W", "E") and text.strip() and not text.startswith("PUPPA-WHO|"):
                self.tail = (self.tail + [text])[-3:]
                m = self.DEVICE.search(text)
                if m:
                    code = m.group(2)
                    copied = copy_text(code)
                    self.msg = f"Logga in på {m.group(1).split('//')[-1]} med koden {code}" + \
                        (" (kopierad)." if copied else ".")
        if not ANIM and kind == "H" and not text.startswith("PUPPA-WHO|"):
            out("    " + dim(text))

    def frame_lines(self, mood, lift=0):
        el = int(time.time() - self.t0)
        msg = self.msg or f"{self.title} … {el}s"
        w = width()
        lines = frame(msg, mood, lift=lift, sub=None if self.msg else "Ctrl+C avbryter")
        with self.lock:
            for t in self.tail:
                for l in wrap(t, w - 8, 2):
                    lines.append("    " + dim(l))
        return lines

    def loop(self):
        i = 0
        while not self.stop.is_set():
            el = time.time() - self.t0
            mood = "blink" if el % 3.1 < 0.18 else ("left" if int(el / 1.3) % 4 == 1 else
                                                     "right" if int(el / 1.3) % 4 == 3 else "wait")
            self.live.draw(self.frame_lines(mood, lift=(i // 4) % 2 if not self.msg else 0))
            i += 1
            self.stop.wait(0.11)

    def __enter__(self):
        if ANIM:
            sys.stdout.write("\033[?25l")
            self.th.start()
        else:
            out(f"  {accent(LOGO)} {self.title} …")
        return self

    def __exit__(self, *a):
        self.stop.set()
        if ANIM:
            self.th.join()
            self.live.clear()
            sys.stdout.write("\033[?25h")
            sys.stdout.flush()


def run_ps(title, code, secret=None, show_output=True, max_lines=None):
    """Kör i arbetsprocessen med väntanimation. -> (ok, antal, rader)."""
    r = Runner(title)
    try:
        with r:
            ok, n = WORKER.run(code, r.event, secret)
    except KeyboardInterrupt:
        out(yellow("  Avbrutet. PowerShell-sessionen startas om (anslutningarna nollställs)."))
        WORKER.stop()
        WORKER.warm()
        return False, 0, r.lines
    if show_output:
        print_output(r.lines, max_lines, quiet_host=not ANIM)
    return ok, n, r.lines


HINTS = [
    (r"AADSTS700016", "ClientId:t finns inte i den här tenanten. Fel tenant, eller så är PnP-appen inte registrerad där."),
    (r"AADSTS65001|consent", "Appen saknar samtycke. En Global Admin måste godkänna den (Entra ID → Appregistreringar → API-behörigheter)."),
    (r"AADSTS50020|AADSTS50034", "Kontot finns inte i kundens tenant. Logga in med ett konto hos kunden."),
    (r"AADSTS7000215|AADSTS700027|certificate", "Certifikatet godtas inte. Kontrollera att rätt certifikat är uppladdat på appen."),
    (r"\(403\)|Forbidden|Access is denied|Åtkomst nekad", "Behörighet saknas. Kontot/appen har inte rätt roll eller rättighet för det här."),
    (r"Request_ResourceNotFound|does not exist|couldn't find|Couldn't find object",
     "Hittade inget med det namnet. Skriv en del av namnet så söker Puppa fram rätt objekt."),
    (r"is not recognized as (a|the) name of a cmdlet", "Ett kommando saknas – modulen är inte installerad eller så är du inte ansluten."),
]


def friendly(errs):
    text = " ".join(errs)
    for rx, msg in HINTS:
        if re.search(rx, text, re.I):
            return msg
    return ""


def print_output(lines, max_lines=None, quiet_host=False):
    global DIRTY
    DIRTY = True
    w = width()
    # Graph-fel bär med sig datum, headrar och request-id – det ingen behöver läsa i terminalen
    lines = [(k, re.split(r"\s+(?:Date|Headers)\s*:", t, 1)[0] if k == "E" else t) for k, t in lines]
    errs = [t for k, t in lines if k == "E"]
    lines = [(k, t) for k, t in lines if not t.startswith("PUPPA-WHO|")]
    lines = [(k, t) for k, t in lines
             if not (k in ("H", "W") and any(e[:50] in t or t.replace("WARNING: ", "")[:50] in e for e in errs if e))]
    body = [t for k, t in lines if k == "O"]
    while body and not body[0].strip():
        body.pop(0)
    while body and not body[-1].strip():
        body.pop()
    if not quiet_host:
        for k, t in lines:
            if k == "H" and t.strip():
                for l in wrap(t, w - 6, 6):
                    out("    " + l)
    for k, t in lines:
        if k == "W":
            for i, l in enumerate(wrap(t, w - 8, 4)):
                out(("  " + yellow("! ") if i == 0 else "    ") + yellow(l))
    limit = max_lines or max(20, 400)
    for l in body[:limit]:
        out("  " + l)
    if len(body) > limit:
        out(dim(f"  … +{len(body) - limit} rader. Välj 'Exportera till Excel' nedan för att få allt."))
    for k, t in lines:
        if k == "E":
            for i, l in enumerate(wrap(t, w - 8, 5)):
                out(("  " + red("✘ ") if i == 0 else "    ") + red(l))
    tip = friendly(errs)
    if tip:
        for i, l in enumerate(wrap(tip, w - 8, 4)):
            out(("  " + yellow("→ ") if i == 0 else "    ") + l)


# ───────────────────────────── kommandovisning ─────────────────────────────
TOK = re.compile(r"('(?:[^']|'')*'|\"[^\"]*\"|#.*$|\$[\w:]+|\b[A-Z][a-z]+-[A-Za-z0-9]+\b|-[A-Za-z]+\b)")


def hl(line):
    if not COLOR:
        return line

    def rep(m):
        t = m.group(0)
        if t.startswith("#"):
            return dim(t)
        if t[0] in "'\"":
            return green(t)
        if t.startswith("$"):
            return cyan(t)
        if t.startswith("-"):
            return dim(t)
        return accent(t)
    return TOK.sub(rep, line)


def show_command(code, title="Kommando som körs"):
    """Hela kommandot syns – långa rader bryts (↪) i stället för att klippas."""
    global DIRTY
    DIRTY = True
    w = width()
    inner = max(16, w - 6)
    out("  " + dim("┌ ") + bold(title) + " " + dim("─" * max(0, min(inner, 70) - len(title) - 1)))
    for raw in code.rstrip().splitlines():
        parts = hard_wrap(raw, inner)
        for k, p in enumerate(parts):
            pre = dim("│ ") if k == 0 else dim("│") + dim("↪")
            out("  " + pre + hl(p))
    out("  " + dim("└" + "─" * max(0, min(inner, 70) + 1)))


# ───────────────────────────── anslutningar ─────────────────────────────
AUTHS = {"interactive": "Interaktiv (webbläsare)", "device": "Enhetskod (device login)",
         "certificate": "Certifikat (app-only)"}
CONN_NAMES = {"pnp": "SharePoint (PnP)", "exo": "Exchange Online", "teams": "Teams", "graph": "Microsoft Graph"}
WHO = "PUPPA-WHO|"
# Efter inloggning: vem och vilken tenant? Puppa sparar svaret i tenant-listan.
WHO_PS = {
    "graph": "$c = Get-MgContext; $o = Get-MgOrganization -ErrorAction SilentlyContinue | Select-Object -First 1\n"
             "Write-Host ('PUPPA-WHO|{0}|{1}|{2}|{3}' -f $c.TenantId, $c.Account, "
             "(($o.VerifiedDomains | Where-Object IsInitial | Select-Object -First 1).Name), $o.DisplayName)",
    "exo": "$ci = Get-ConnectionInformation | Select-Object -Last 1; $oc = Get-OrganizationConfig -ErrorAction SilentlyContinue\n"
           "Write-Host ('PUPPA-WHO|{0}|{1}|{2}|{3}' -f $ci.TenantID, $ci.UserPrincipalName, $oc.Name, $oc.DisplayName)",
    "teams": "Write-Host ('PUPPA-WHO|{0}|{1}|{2}|' -f $t.TenantId, \"$($t.Account)\", $t.TenantDomain)",
}


def keep_logins():
    """Spara inloggningar mellan körningar (tokens i nyckelringen / MSAL-cachen)?
    Av som standard på macOS: en osignerad pwsh (Homebrew) får annars en nyckelringsfråga vid varje inloggning."""
    v = config().get("keep_logins")
    return (not IS_MAC) if v is None else bool(v)


def pnp_connect_code(t, url):
    base = f"Connect-PnPOnline -Url {psq(url)} -ClientId {psq(t['client_id'])} -Tenant {psq(t['tenant'])}"
    a = t.get("auth", "interactive")
    if a == "certificate":
        return base + f" -CertificatePath {psq(t['cert_path'])} -CertificatePassword {SECRET_EXPR} -ErrorAction Stop"
    if a == "device":
        return base + " -DeviceLogin" + (" -PersistLogin" if keep_logins() else "") + " -ErrorAction Stop"
    return base + " -Interactive" + (" -PersistLogin" if keep_logins() else "") + " -ErrorAction Stop"


def connect_code(kind, t, url=None, scopes=None):
    """t = vald tenant eller None (då väljer du konto/tenant i inloggningsrutan)."""
    dev = bool(t) and t.get("auth") == "device"
    tid = (t.get("tenant_id") or t.get("tenant")) if t else ""
    if kind == "pnp":
        return pnp_connect_code(t, url or t["admin_url"])
    if kind == "exo":
        s = "Connect-ExchangeOnline -ShowBanner:$false"
        if t and t.get("upn"):
            s += f" -UserPrincipalName {psq(t['upn'])}"
        return s + (" -Device" if dev else "") + " -ErrorAction Stop"
    if kind == "teams":
        s = "$t = Connect-MicrosoftTeams" + (f" -TenantId {psq(tid)}" if tid else "")
        if t and t.get("upn"):
            s += f" -AccountId {psq(t['upn'])}"
        return s + (" -UseDeviceAuthentication" if dev else "") + " -ErrorAction Stop"
    if kind == "graph":
        sc = ",".join(psq(x) for x in sorted(set(scopes or []) | {"User.Read"}))
        return "Connect-MgGraph" + (f" -TenantId {psq(tid)}" if tid else "") + f" -Scopes {sc} -NoWelcome" + \
               ("" if keep_logins() else " -ContextScope Process") + (" -UseDeviceCode" if dev else "") + " -ErrorAction Stop"
    raise ValueError(kind)


def need_secret(kind, t):
    if kind == "pnp" and t.get("auth") == "certificate":
        s = sec_get(t)
        if not s:
            raise RuntimeError(f"Hittar inget certifikatlösenord för {t['name']} i valvet {VAULT}.")
        return s
    return None


def is_connected(kind, t, url=None, scopes=None):
    cur = WORKER.conn.get(kind)
    if not cur or not t or cur[0] != t["name"]:
        return False
    if kind == "pnp":
        return cur[1].rstrip("/").lower() == (url or t["admin_url"]).rstrip("/").lower()
    if kind == "graph":
        return set(scopes or []) <= set(cur[1])
    return True


def ensure(kind, t, url=None, scopes=None, show=True):
    """Logga in om det behövs. -> tenanten (kan vara en ny, sparad efter inloggningen) eller None."""
    if kind == "pnp":
        t = pnp_ready(t)
        if not t:
            return None
    if is_connected(kind, t, url, scopes):
        return t
    if kind == "graph" and t and WORKER.conn.get("graph") and WORKER.conn["graph"][0] == t["name"]:
        scopes = sorted(set(scopes or []) | set(WORKER.conn["graph"][1]))
    code = connect_code(kind, t, url, scopes)
    if show:
        show_command(code, "Loggar in")
    if DRY:
        return t or {}
    try:
        secret = need_secret(kind, t)
    except RuntimeError as e:
        say(str(e), "sad")
        return None
    full = code + ("\n" + WHO_PS[kind] if kind in WHO_PS else "")
    title = f"Loggar in på {CONN_NAMES[kind]}" + (f" ({t['name']})" if t else " – välj konto i webbläsaren")
    ok, _, lines = run_ps(title, full, secret, show_output=False)
    if not ok:
        print_output(lines, quiet_host=not ANIM)
        say(f"Kunde inte logga in på {CONN_NAMES[kind]}. Felet står ovanför.", "sad")
        return None
    who = next((x[1][len(WHO):].split("|") for x in lines if x[0] == "H" and x[1].startswith(WHO)), None)
    if who and any(who):
        before = t["name"] if t else None
        t, new = remember_login(who)
        if before and before != t["name"]:
            out("  " + yellow("!") + f" Du loggade in i {bold(t['name'])}, inte {before}. Bytte tenant.")
        if new:
            out("  " + green("✔") + f" Ny tenant i listan: {bold(t['name'])} " + dim(t.get("tenant", "")))
    if not t:
        return None
    WORKER.conn[kind] = (t["name"], url or t["admin_url"]) if kind == "pnp" else \
        (t["name"], sorted(scopes or [])) if kind == "graph" else (t["name"],)
    out("  " + green("✔") + f" Inloggad på {CONN_NAMES[kind]}" + dim(f" · {t['name']}" +
                                                                    (f" · {t['upn']}" if t.get("upn") else "")))
    return t


def conn_status():
    t = active_tenant()
    w = width()
    if not t:
        head = f"  {dim('Tenant')} {dim('ingen vald – jag loggar in när ett recept behöver det')}"
    else:
        head = f"  {dim('Tenant')} {bold(t['name'])}"
        if w >= 58:
            head += dim(f"  {t.get('tenant', '')}")
        if w >= 96 and t.get("upn"):
            head += dim(f"  ·  {t['upn']}")
    parts = []
    for k, short in (("pnp", "PnP"), ("exo", "EXO"), ("teams", "Teams"), ("graph", "Graph")):
        on = bool(t) and WORKER.conn.get(k, ("",))[0] == t["name"]
        parts.append((green("● ") if on else dim("○ ")) + (short if on else dim(short)))
    return [head, "  " + "  ".join(parts)]


# ───────────────────────────── uppslag (sök i tenanten) ─────────────────────────────
# I ett recept: param: Anvandare | Användare | | @user      (@typ+ = flera värden, kommaseparerade)
# Varje uppslag skriver "PUPPA-PICK|värde|namn|detalj" per träff. $Q är det du skrev.
PICK = "PUPPA-PICK|"
PNP_APPS_PS = r"""
$sp = '00000003-0000-0ff1-ce00-000000000000'
$pnp = Get-MgApplication -All -Property 'id,appId,displayName,createdDateTime,requiredResourceAccess,publicClient' |
    Where-Object { $_.DisplayName -match 'PnP' -or (($_.RequiredResourceAccess.ResourceAppId -contains $sp) -and
                   (@($_.PublicClient.RedirectUris) -match 'localhost')) } | Sort-Object CreatedDateTime -Descending
"""
LOOKUPS = {
    "user": ("graph:User.Read.All", "användare", r"""
$s = $Q -replace '"', ''
$p = @{ Top = 40; Property = 'displayName,userPrincipalName,accountEnabled,userType'; ConsistencyLevel = 'eventual' }
if ($s) { $p.Search = ('"displayName:{0}" OR "userPrincipalName:{0}" OR "mail:{0}"' -f $s) }
Get-MgUser @p | ForEach-Object { Write-Host ('PUPPA-PICK|{0}|{1}|{2}' -f $_.UserPrincipalName, $_.DisplayName,
    ($_.UserPrincipalName + $(if ($_.UserType -eq 'Guest') { ' · gäst' }) + $(if (-not $_.AccountEnabled) { ' · avstängd' }))) }
"""),
    "group": ("graph:Group.Read.All", "grupper", r"""
$s = $Q -replace '"', ''
$p = @{ Top = 40; Property = 'displayName,mail,groupTypes,securityEnabled'; ConsistencyLevel = 'eventual' }
if ($s) { $p.Search = ('"displayName:{0}" OR "mail:{0}"' -f $s) }
Get-MgGroup @p | ForEach-Object {
    $typ = if ($_.GroupTypes -contains 'Unified') { 'Microsoft 365' } elseif ($_.SecurityEnabled) { 'säkerhet' } else { 'distribution' }
    Write-Host ('PUPPA-PICK|{0}|{1}|{2}' -f $(if ($_.Mail) { $_.Mail } else { $_.DisplayName }), $_.DisplayName, "$($_.Mail) · $typ") }
"""),
    "sku": ("graph:Organization.Read.All", "licenser", r"""
Get-MgSubscribedSku -All | Where-Object { $_.PrepaidUnits.Enabled -gt 0 } | ForEach-Object {
    Write-Host ('PUPPA-PICK|{0}|{0}|{1} av {2} använda' -f $_.SkuPartNumber, $_.ConsumedUnits, $_.PrepaidUnits.Enabled) }
"""),
    "mailbox": ("exo", "postlådor", r"""
$t = 'UserMailbox','SharedMailbox','RoomMailbox','EquipmentMailbox'
$r = if ($Q) { Get-EXORecipient -Anr $Q -RecipientTypeDetails $t -ResultSize 40 } else { Get-EXORecipient -RecipientTypeDetails $t -ResultSize 40 }
$r | ForEach-Object { Write-Host ('PUPPA-PICK|{0}|{1}|{0} · {2}' -f $_.PrimarySmtpAddress, $_.DisplayName, $_.RecipientTypeDetails) }
"""),
    "recipient": ("exo", "mottagare", r"""
$r = if ($Q) { Get-EXORecipient -Anr $Q -ResultSize 40 } else { Get-EXORecipient -ResultSize 40 }
$r | Where-Object PrimarySmtpAddress | ForEach-Object { Write-Host ('PUPPA-PICK|{0}|{1}|{0} · {2}' -f $_.PrimarySmtpAddress, $_.DisplayName, $_.RecipientTypeDetails) }
"""),
    "dl": ("exo", "distributionslistor", r"""
$t = 'MailUniversalDistributionGroup','MailUniversalSecurityGroup','RoomList'
$r = if ($Q) { Get-EXORecipient -Anr $Q -RecipientTypeDetails $t -ResultSize 40 } else { Get-EXORecipient -RecipientTypeDetails $t -ResultSize 40 }
$r | ForEach-Object { Write-Host ('PUPPA-PICK|{0}|{1}|{0} · {2}' -f $_.PrimarySmtpAddress, $_.DisplayName, $_.RecipientTypeDetails) }
"""),
    "site": ("pnp:admin", "sajter", r"""
Get-PnPTenantSite | ForEach-Object { Write-Host ('PUPPA-PICK|{0}|{1}|{0}' -f $_.Url, $(if ($_.Title) { $_.Title } else { $_.Url })) }
"""),
    "spuser": ("pnp:admin", "användare", r"""
$s = $Q -replace "'", "''"
$f = if ($s) { "startswith(displayName,'$s') or startswith(userPrincipalName,'$s') or startswith(mail,'$s') or startswith(surname,'$s')" } else { $null }
$p = @{ EndIndex = 40 }; if ($f) { $p.Filter = $f }
Get-PnPEntraIDUser @p | ForEach-Object { Write-Host ('PUPPA-PICK|{0}|{1}|{0}' -f $_.UserPrincipalName, $_.DisplayName) }
"""),
    "team": ("teams", "team", r"""
Get-Team | ForEach-Object { Write-Host ('PUPPA-PICK|{0}|{0}|{1}' -f $_.DisplayName, "$($_.Visibility)$(if ($_.Archived) { ' · arkiverat' })") }
"""),
    "app": ("graph:Application.Read.All", "appregistreringar", r"""
$s = $Q -replace '"', ''
$p = @{ Top = 40; Property = 'id,appId,displayName,createdDateTime'; ConsistencyLevel = 'eventual' }
if ($s) { $p.Search = ('"displayName:{0}"' -f $s) }
Get-MgApplication @p | ForEach-Object { Write-Host ('PUPPA-PICK|{0}|{1}|{0} · skapad {2:yyyy-MM-dd}' -f $_.AppId, $_.DisplayName, $_.CreatedDateTime) }
"""),
    # PnP-appar: namnet innehåller PnP, eller en publik klient (http://localhost) med SharePoint-rättigheter
    "pnpapp": ("graph:Application.Read.All", "PnP-appar", PNP_APPS_PS + r"""
$pnp | ForEach-Object { Write-Host ('PUPPA-PICK|{0}|{1}|{0} · skapad {2:yyyy-MM-dd}' -f $_.AppId, $_.DisplayName, $_.CreatedDateTime) }
"""),
}
LIST_ALL = {"site", "team", "sku", "pnpapp"}          # hämtas en gång per session och filtreras här i Puppa
LOOKUP_CACHE = {}


def need_for(spec):
    """'graph:User.Read.All' -> ('graph', None, ['User.Read.All']) · 'pnp:admin' -> ('pnp', admin_url-markör, None)"""
    kind, _, arg = spec.partition(":")
    if kind == "graph":
        return kind, None, [s for s in arg.split(",") if s]
    return kind, None, None


def run_lookup(kind, t, query):
    """-> (t, [(värde, namn, detalj)]) eller (t, None) om uppslaget inte gick."""
    spec, _, code = LOOKUPS[kind]
    ck, _, scopes = need_for(spec)
    t2 = ensure(ck, t, None, scopes, show=False)
    if not t2:
        return t, None
    t = t2
    cache_key = (kind, t["name"])
    if kind in LIST_ALL and cache_key in LOOKUP_CACHE:
        items = LOOKUP_CACHE[cache_key]
    else:
        q = "" if kind in LIST_ALL else query
        ok, _, lines = run_ps(f"Söker {LOOKUPS[kind][1]}" + (f" ”{query}”" if q else ""),
                              f"$Q = {psq(q)}\n{code}", show_output=False)
        items = []
        for k, txt in lines:
            if k == "H" and txt.startswith(PICK):
                f = (txt[len(PICK):].split("|") + ["", "", ""])[:3]
                if f[0]:
                    items.append(tuple(f))
        if not ok and not items:
            print_output(lines, quiet_host=True)
            return t, None
        if kind in LIST_ALL:
            LOOKUP_CACHE[cache_key] = items
    ranked = fuzzy_rank(query, items, lambda x: f"{x[1]} {x[0]} {x[2]}") if query else items
    return t, ranked


def lookup_param(prm, t, default):
    """Fråga efter en del av namnet, sök i tenanten och låt användaren välja. -> (t, värde)"""
    kind = prm.lookup.rstrip("+")
    many = prm.lookup.endswith("+")
    chosen = []
    q_default = "" if kind == "site" else default        # sajt: sök på namn, inte på en halv URL
    while True:
        what = prm.question + (" – fler: lägg till en i taget" if many and chosen else "")
        q = ask_text(what + (" (valfritt)" if prm.optional else "") + " · skriv en del av namnet",
                     q_default, required=not prm.optional and not chosen)
        q_default = ""
        if not q:
            break
        if "@" in q and " " not in q and kind not in ("site", "team", "sku"):
            v = q                                            # fullständig adress: använd direkt
        elif kind == "site" and re.match(r"https?://", q):
            v = q                                            # inklistrad länk (sajt, bibliotek eller mapp)
        else:
            t, hits = run_lookup(kind, t, q)
            if hits is None:
                out("    " + yellow(f"Kunde inte söka – använder ”{q}” som det är."))
                v = q
            else:
                exact = [h for h in hits if h[0].lower() == q.lower() or h[1].lower() == q.lower()]
                if len(exact) == 1 or len(hits) == 1:
                    h = (exact or hits)[0]
                    v = h[0]
                    out("    " + green("→ ") + bold(h[1]) + dim(f"  {h[0]}" if h[0] != h[1] else ""))
                elif not hits:
                    out("    " + yellow(f"Hittade inga {LOOKUPS[kind][1]} som matchar ”{q}”."))
                    v = choose([dict(label="Sök igen", value="__again"),
                                dict(label=f"Använd ”{q}” som det är", value=q)], title="Vad vill du göra?")
                    if v is None:
                        raise Cancel()
                    if v == "__again":
                        continue
                else:
                    rows = [dict(label=h[1], hint=h[2] or h[0], value=h[0]) for h in hits[:200]]
                    rows += [("head", ""), dict(label="Sök igen…", value="__again", key="s"),
                             dict(label=f"Använd ”{q}” som det är", value="__raw", key="r")]
                    v = choose(rows, title=f"{len(hits)} {LOOKUPS[kind][1]} matchar ”{q}”" +
                               (" – skriv för att filtrera" if len(hits) > 9 else ""))
                    if v is None:
                        raise Cancel()
                    if v == "__again":
                        continue
                    if v == "__raw":
                        v = q
        chosen.append(v)
        if not many or not confirm("Lägg till fler?", False):
            break
    return t, ",".join(chosen)


# ───────────────────────────── recept ─────────────────────────────
class Param:
    def __init__(self, spec):
        f = [x.strip() for x in spec.split("|")]
        name = f[0]
        self.optional = name.endswith("?")
        self.name = name.rstrip("?")
        if not re.match(r"^[A-Za-z_]\w*$", self.name):
            raise ValueError(f"ogiltigt parameternamn '{self.name}' (bara bokstäver a–z, siffror, _)")
        self.question = f[1] if len(f) > 1 and f[1] else self.name
        self.default = f[2] if len(f) > 2 else ""
        spec = f[3] if len(f) > 3 else ""
        self.lookup = spec.strip() if spec.strip().startswith("@") else ""
        if self.lookup:
            self.lookup = self.lookup[1:]
            if self.lookup.rstrip("+") not in LOOKUPS:
                raise ValueError(f"okänt uppslag '@{self.lookup}' (finns: {', '.join('@' + k for k in LOOKUPS)})")
        self.choices = [] if self.lookup else [x.strip() for x in spec.split(",") if x.strip()]


class Recipe:
    def __init__(self, path):
        self.path = path
        self.id = path.stem
        text = path.read_text(encoding="utf-8")
        m = re.match(r"\s*<#\s*PUPPA\s*\n(.*?)#>[ \t]*\n?", text, re.S)
        meta, self.params = {}, []
        if m:
            for line in m.group(1).splitlines():
                if ":" not in line or line.strip().startswith("#"):
                    continue
                k, _, v = line.partition(":")
                k = k.strip().lower().replace("ä", "a")
                if k == "param":
                    self.params.append(Param(v.strip()))
                else:
                    meta[k] = v.strip()
            self.body = text[m.end():].strip("\n")
        else:
            self.body = text.strip("\n")
        self.name = meta.get("namn", self.id)
        self.category = meta.get("kategori", "Övrigt")
        self.kind = meta.get("typ", "läs").lower()
        self.needs = meta.get("kraver", "ingen").strip()
        self.desc = meta.get("beskrivning", "")

    @property
    def writes(self):
        return self.kind.startswith("änd") or self.kind.startswith("and") or self.kind == "skriv"

    def badge(self):
        if self.writes:
            return yellow("ändrar")
        if self.kind.startswith("lok"):
            return cyan("lokal")
        return green("läs")

    def needs_list(self):
        """'pnp-admin, graph:A,B' -> [('pnp', 'admin'), ('graph', 'A,B')]"""
        res = []
        for part in re.split(r"[;\s]+(?=[a-z])", self.needs):
            part = part.strip().strip(",")
            if not part or part == "ingen":
                continue
            if part.startswith("pnp"):
                res.append(("pnp", part.split(":", 1)[1] if ":" in part else ("admin" if "admin" in part else "root")))
            elif part.startswith("graph"):
                res.append(("graph", part.split(":", 1)[1] if ":" in part else "User.Read.All"))
            else:
                res.append((part, ""))
        return res


def load_recipes():
    rs = {}
    for d in (RECIPES, USER_RECIPES):
        if d.exists():
            for f in sorted(d.glob("*.ps1")):
                try:
                    r = Recipe(f)
                    rs[r.id] = r
                except Exception as e:
                    out(red(f"  Kunde inte läsa {f.name}: {e}"))
    order = ["SharePoint", "OneDrive", "Exchange", "Grupper", "Användare", "Teams", "Entra ID", "Excel", "Lokalt"]
    return sorted(rs.values(), key=lambda r: (order.index(r.category) if r.category in order else 50,
                                              r.category, r.writes, r.name))


def find_recipe(name, rs=None):
    rs = rs or load_recipes()
    n = name.lower()
    for r in rs:
        if r.id.lower() == n or r.name.lower() == n:
            return r
    hits = [r for r in rs if n in r.id.lower() or n in r.name.lower()]
    return hits[0] if len(hits) == 1 else None


def placeholders(t):
    t = t or {}
    name = t.get("name") or "puppa"
    return {"tenant": t.get("tenant", ""), "prefix": t.get("prefix", ""), "adminUrl": t.get("admin_url", ""),
            "rootUrl": t.get("root_url", ""), "upn": t.get("upn", ""), "kund": name, "profil": name,
            "datum": datetime.now().strftime("%Y-%m-%d")}


def fill(s, ph):
    return re.sub(r"\{(\w+)\}", lambda m: ph.get(m.group(1), m.group(0)), s)


def collect_params(r, p, given):
    """-> (värden, tenant). Uppslag (@user, @group …) kan logga in och därmed sätta tenanten."""
    ph = placeholders(p)
    vals = {}
    for prm in r.params:
        default = fill(prm.default, ph)
        if prm.name in given:
            vals[prm.name] = given[prm.name]
            continue
        if not ITTY:
            if not default and not prm.optional:
                raise Cancel(f"saknar {prm.name}=…")
            vals[prm.name] = default
            continue
        if prm.lookup and not DRY:
            p, v = lookup_param(prm, p, default)
            vals[prm.name] = v
            continue
        if prm.choices:
            rows = [dict(label=ch, value=ch) for ch in prm.choices]
            v = choose(rows, title=prm.question, start=default or None)
            if v is None:
                raise Cancel()
        else:
            v = ask_text(prm.question + ("" if not prm.optional else " (valfritt)"), default,
                         required=not prm.optional)
        vals[prm.name] = v
    return vals, p


def build_script(r, vals, p=None):
    head = "\n".join(f"${k} = {psq(v)}" for k, v in vals.items())
    return (head + "\n" if head else "") + fill(r.body, placeholders(p))


def run_recipe(r, given=None, yes=False):
    t = active_tenant()
    needs = r.needs_list()
    w = width()
    out(f"  {accent('▍')}{bold(r.name)}  {r.badge()}" + (dim(f"  {t['name']}") if t and needs else ""))
    for l in wrap(r.desc, w - 6, 3) if r.desc else []:
        out("    " + dim(l))
    print()
    if any(k == "pnp" for k, _ in needs):
        t = pnp_ready(t)
        if not t:
            return False
    try:
        vals, t = collect_params(r, t, given or {})
    except Cancel as e:
        if e.args:
            out(red("  " + e.args[0]))
        out(dim("  Avbrutet."))
        return False
    pre = []
    for kind, arg in needs:
        url = scopes = None
        if kind == "pnp":
            url = t["admin_url"] if arg == "admin" else t["root_url"] if arg == "root" else vals.get(arg, arg)
        if kind == "graph":
            scopes = [s.strip() for s in arg.split(",") if s.strip()]
        pre.append((kind, url, scopes))
    print()
    if DRY:
        for kind, url, scopes in pre:
            if not is_connected(kind, t, url, scopes):
                show_command(connect_code(kind, t, url, scopes), "Först: logga in")
        show_command(build_script(r, vals, t))
        say("Torrkörning: inget kördes. Så här skulle det se ut.", "happy")
        return True
    # Logga in först (det ändrar inget) – då vet vi vilken tenant ändringen hamnar i innan du bekräftar.
    for kind, url, scopes in pre:
        t2 = ensure(kind, t, url, scopes)
        if not t2:
            return False
        t = t2
    script = build_script(r, vals, t)
    show_command(script)
    if r.writes and not yes:
        print()
        say(f"Det här ändrar i {t['name']} ({t.get('tenant', '')}). Kolla kommandot ovan först." if t else
            "Det här ändrar något. Kolla kommandot ovan först.", "worry")
        if not confirm("Kör på riktigt?", False):
            say("Okej, inget ändrat.", "happy")
            return False
    ok, n, _ = run_ps(f"Kör {r.name}", script)
    print()
    if ok:
        word = f"{n} rader" if n != 1 else "1 rad"
        if r.writes:
            celebrate("Klart! Ändringen är gjord.")
        else:
            say(f"Klart! {word}." if n else "Klart!", "proud")
            if n > 1 and r.category != "Excel" and ITTY and not DRY:
                offer_export(r, t)
    else:
        say("Det gick inte hela vägen. Felet står ovanför i rött.", "sad")
    return ok


def offer_export(r, t):
    """Efter ett läsrecept med flera rader: exportera direkt, eller tillbaka."""
    global DIRTY
    try:
        v = choose([dict(label="Tillbaka till menyn", value="back"),
                    dict(label="Exportera till Excel", hint="formaterad .xlsx i ~/Documents/Puppa", value="excel-exportera", key="x"),
                    dict(label="Exportera till CSV", hint="semikolon, UTF-8", value="excel-csv", key="c")],
                   title="Vad vill du göra med resultatet?", footer="↑/↓ · enter · x Excel · c CSV · q tillbaka")
    except Cancel:
        return
    if v in (None, "back"):
        DIRTY = False            # användaren har redan sett resultatet – ingen extra paus
        return
    ex = find_recipe(v)
    if ex:
        name = slug(f"{(t or {}).get('name', 'puppa')}-{r.id}-{datetime.now():%Y-%m-%d}")
        run_recipe(ex, given={"Filnamn": name, "Oppna": "Ja"}, yes=True)


RECIPE_TEMPLATE = """<#PUPPA
namn:        {name}
kategori:    {cat}
typ:         {kind}
kräver:      {needs}
beskrivning: Beskriv kort vad receptet gör.
param:       Namn | Vad ska Puppa fråga om? | standardvärde
param:       Valfri? | Ett valfritt värde (tomt går bra) |
#>
# Parametrarna finns som variabler: $Namn, $Valfri.
# Platshållare i standardvärden och skript: {{tenant}} {{prefix}} {{adminUrl}} {{rootUrl}} {{upn}} {{kund}} {{datum}}
# Returnera objekt (Select-Object …) så kan resultatet exporteras till Excel efteråt.
Write-Host "Hej $Namn"
"""


def new_recipe():
    print()
    say("Ett recept är en vanlig .ps1 med ett litet huvud överst. Jag skapar en mall åt dig.", "happy")
    try:
        name = ask_text("Vad ska receptet heta?", required=True)
        cat = choose([dict(label=x, value=x) for x in ["SharePoint", "OneDrive", "Exchange", "Grupper", "Användare",
                                                      "Teams", "Entra ID", "Excel", "Lokalt"]],
                     title="Kategori")
        if cat is None:
            raise Cancel()
        kind = choose([dict(label="läs", hint="körs direkt", value="läs"),
                       dict(label="ändrar", hint="kräver bekräftelse", value="ändrar"),
                       dict(label="lokal", hint="rör inte tenanten", value="lokal")], title="Typ")
        if kind is None:
            raise Cancel()
        needs = choose([dict(label="ingen", value="ingen"), dict(label="pnp-admin", hint="SharePoint-admin", value="pnp-admin"),
                        dict(label="pnp:Url", hint="sajt från parametern Url", value="pnp:Url"),
                        dict(label="exo", value="exo"), dict(label="teams", value="teams"),
                        dict(label="graph:User.Read.All", value="graph:User.Read.All")], title="Kräver anslutning")
        if needs is None:
            raise Cancel()
    except Cancel:
        out(dim("  Avbrutet."))
        return
    USER_RECIPES.mkdir(parents=True, exist_ok=True)
    f = USER_RECIPES / (slug(name) + ".ps1")
    if f.exists() and not confirm(f"{f.name} finns redan. Skriva över?", False):
        return
    f.write_text(RECIPE_TEMPLATE.format(name=name, cat=cat, kind=kind, needs=needs), encoding="utf-8")
    say(f"Mallen ligger i {f}. Redigera den, så dyker receptet upp i menyn direkt.", "proud")
    if confirm("Öppna i editorn?", True):
        open_file(f)


# ───────────────────────────── tenants & PnP-appregistrering ─────────────────────────────
PERMS = {
    "read": ("Bara läsa", "Sites.Read.All · Group.Read.All · User.Read.All",
             "-SharePointApplicationPermissions Sites.Read.All -GraphApplicationPermissions Group.Read.All,User.Read.All"),
    "full": ("Läsa och ändra", "Sites.FullControl.All · Group.ReadWrite.All · User.Read.All",
             "-SharePointApplicationPermissions Sites.FullControl.All -GraphApplicationPermissions Group.ReadWrite.All,User.Read.All"),
    "pnp": ("PnP:s standard", "det PnP.PowerShell väljer själv", ""),
}


def check_tenant(v):
    if " " in v.strip():
        return "Inga mellanslag. T.ex. contoso, contoso.onmicrosoft.com eller contoso.se."
    return None


def ago(iso):
    try:
        s = (datetime.now() - datetime.fromisoformat(iso)).total_seconds()
    except Exception:
        return ""
    if s < 3600:
        return "nyss"
    if s < 86400:
        return f"{int(s // 3600)} h sedan"
    return f"{int(s // 86400)} d sedan"


def tenant_rows():
    act = config().get("active")
    w = width()
    ts = sorted(tenants(), key=lambda t: t.get("last") or t.get("created") or "", reverse=True)
    rows = []
    for t in ts:
        hint = t.get("tenant", "")
        if w >= 96 and t.get("upn"):
            hint += "  · " + t["upn"]
        if w >= 120 and t.get("last"):
            hint += "  · " + ago(t["last"])
        badge = green("aktiv") if t["name"] == act else (cyan("PnP") if t.get("client_id") else "")
        rows.append(dict(label=t["name"], hint=hint, value=t["name"], badge=badge))
    return rows


def add_tenant():
    """Lägg till en tenant för hand (oftast behövs det inte – inloggning sparar tenanten automatiskt)."""
    traw = ask_text("Tenantens domän? (contoso, contoso.onmicrosoft.com eller contoso.se)", required=True,
                    check=check_tenant)
    domain, prefix = tenant_info(traw)
    ts = tenants()
    if any(x.get("tenant", "").lower() == domain for x in ts):
        say(f"{domain} finns redan i listan.", "happy")
        t = next(x for x in ts if x.get("tenant", "").lower() == domain)
        set_config(active=t["name"])
        return t
    base = (prefix or domain.split(".")[0]).capitalize()
    name = ask_text("Vad ska den heta i listan?", unique_name(base, ts))
    upn = ask_text("Ditt admin-konto där? (valfritt, föreslås vid inloggning)", "")
    t = dict(name=unique_name(name, ts), tenant=domain, tenant_id="", upn=upn, client_id="", auth="interactive",
             created=datetime.now().isoformat(timespec="seconds"))
    t.update(sp_fields(prefix))
    ts.append(t)
    save_tenants(ts)
    use_tenant(t, quiet=True)
    say(f"{t['name']} är tillagd och vald.", "proud")
    return t


def pnp_ready(t):
    """SharePoint-recept går via PnP PowerShell, som kräver en Entra ID-app i tenanten. Saknas den: erbjud registrering."""
    if t and t.get("client_id") and t.get("admin_url"):
        return t
    if not ITTY:
        say("SharePoint-recept behöver en PnP-app för tenanten. Kör: puppa pnp", "worry")
        return None
    where = f"i {t['name']}" if t else "i tenanten"
    say(f"SharePoint-recept går via PnP PowerShell, som behöver en egen app-registrering {where}. "
        "Finns det redan en kan jag söka fram den.", "worry")
    how = choose([dict(label="Sök fram befintlig PnP-app", hint="letar bland appregistreringarna via Graph", value="find"),
                  dict(label="Klistra in ClientId", hint="om du vet vilken app det är", value="paste"),
                  dict(label="Ny PnP-appregistrering…", hint="Register-PnPEntraID… (frågar först)", value="new"),
                  dict(label="Avbryt", value=None, key="q")],
                 title="PnP-app " + where)
    if how is None:
        return None
    if how == "new":
        return pnp_setup(t, intro=False)
    try:
        cid = None
        if how == "find":
            t2, cid = find_pnp_app(t)
            t = t2 or t
            if not cid:
                out(dim("  Du kan klistra in ClientId i stället (esc avbryter)."))
        if not cid:
            cid = ask_text("ClientId (Application ID, en GUID)", required=True,
                           check=lambda v: None if GUID.match(v) else "Det ser inte ut som en GUID.")
        if not t:
            say("Jag vet inte vilken tenant appen hör till – logga in först (Logga in… → Microsoft Graph).", "worry")
            return None
        t = dict(t, client_id=cid)
        t.setdefault("auth", "interactive")
        if not t.get("admin_url"):
            dom, pre = tenant_info(t.get("tenant", ""))
            if not pre:
                pre = ask_text("SharePoint-namnet? (det som står före .sharepoint.com)", required=True, check=check_tenant)
            t.update(sp_fields(pre.lower().replace(".sharepoint.com", "").replace("-admin", "").replace("https://", "")))
    except Cancel:
        out(dim("  Avbrutet. Inget sparat."))
        return None
    save_tenants([x for x in tenants() if x["name"] != t["name"]] + [t])
    use_tenant(t, quiet=True)
    out("  " + green("✔") + f" PnP-appen sparad för {t['name']} · ClientId {cid} · {AUTHS.get(t['auth'], '')}")
    return t


def register_code(app, tenant, auth, perm, outdir):
    dev = " -DeviceLogin" if auth == "device" else ""
    if auth == "certificate":
        extra = PERMS[perm][2]
        return (f"Register-PnPEntraIDApp -ApplicationName {psq(app)} -Tenant {psq(tenant)} "
                f"-OutPath {psq(str(outdir))} -CertificatePassword {SECRET_EXPR} -ValidYears 2"
                + (" " + extra if extra else "") + dev)
    return f"Register-PnPEntraIDAppForInteractiveLogin -ApplicationName {psq(app)} -Tenant {psq(tenant)}" + dev


REG_WRAP = r"""
$ErrorActionPreference = 'Stop'
$r = {cmd}
$h = @{{}}
if ($r -is [System.Collections.IDictionary]) {{ foreach ($k in $r.Keys) {{ $h["$k"] = "$($r[$k])" }} }}
else {{ foreach ($pp in @($r)[0].PSObject.Properties) {{ $h[$pp.Name] = "$($pp.Value)" }} }}
$h | ConvertTo-Json | Set-Content -LiteralPath $env:PUPPA_OUT -Encoding utf8
"""


def pnp_setup(t=None, intro=True):
    """Ny PnP-appregistrering för en tenant: tenant → SharePoint-adress → inloggningssätt → app → klart."""
    print()
    if intro:
        say("Ny PnP-appregistrering. Jag frågar, du svarar – och inget skapas i tenanten utan att du säger ja.",
            "happy", sub="Ctrl+C avbryter när som helst")
    try:
        if not t:
            ts = tenants()
            if ts:
                v = choose(tenant_rows() + [("head", ""), dict(label="En annan tenant…", value="__new", key="n")],
                           title="Vilken tenant?")
                if v is None:
                    raise Cancel()
                t = add_tenant() if v == "__new" else find_tenant(v)
            else:
                t = add_tenant()
        t = dict(t)
        if not t.get("prefix"):
            dom, pre = tenant_info(t.get("tenant", ""))
            if not pre:
                pre = ask_text("SharePoint-namnet? (det som står före .sharepoint.com)", required=True,
                               check=check_tenant)
                pre = pre.lower().replace(".sharepoint.com", "").replace("-admin", "").replace("https://", "")
            t.update(sp_fields(pre))
        t["admin_url"] = ask_text("SharePoint-admin-URL", t["admin_url"]).rstrip("/")
        auth = choose([dict(label=AUTHS["interactive"], hint="du loggar in, delegerade rättigheter", value="interactive"),
                       dict(label=AUTHS["device"], hint="kod i webbläsaren, bra över SSH", value="device"),
                       dict(label=AUTHS["certificate"], hint="obevakat, certifikat + lösenord i valvet", value="certificate")],
                      title="Hur ska PnP logga in?")
        if auth is None:
            raise Cancel()
        t["auth"] = auth
        how = choose([dict(label="Registrera en ny app åt mig", hint="Register-PnPEntraID… (frågar först)", value="new"),
                      dict(label="Det finns redan en app", hint="sök fram den eller klistra in ClientId", value="have")],
                     title=f"PnP-app i {t['name']}")
        if how is None:
            raise Cancel()
        if how == "have":
            t2, cid = ask_client_id(t)
            t = dict(t2, **{k: t[k] for k in ("admin_url", "auth", "prefix", "root_url") if k in t}, client_id=cid)
            if auth == "certificate" and not import_cert(t):
                raise Cancel()
        elif not register_app(t):
            return None
    except Cancel:
        out(dim("  Avbrutet. Inget sparat."))
        return None
    ts = [x for x in tenants() if x["name"] != t["name"]] + [t]
    save_tenants(ts)
    use_tenant(t, quiet=True)
    print()
    celebrate(f"PnP är klart för {t['name']}!")
    print_tenant(t)
    if confirm("Testa anslutningen nu?", True):
        test_pnp(t)
    return t


def find_pnp_app(t):
    """Sök i tenanten (Graph) efter appar som ser ut som PnP-appar och låt användaren välja. -> (t, ClientId|None)"""
    LOOKUP_CACHE.pop(("pnpapp", (t or {}).get("name")), None)
    t2, hits = run_lookup("pnpapp", t, "")
    t = t2 or t or {}
    if hits is None:
        say("Kunde inte söka bland appregistreringarna (kräver Application.Read.All i Graph).", "sad")
        return t, None
    if not hits:
        say(f"Hittade ingen PnP-app i {t.get('name', 'tenanten')} – varken på namn eller som publik klient med SharePoint-rättigheter.",
            "worry")
        return t, None
    rows = [dict(label=h[1], hint=h[2], value=h[0]) for h in hits]
    rows += [("head", ""), dict(label="Ingen av dem", value="__none", key="q")]
    v = choose(rows, title=f"{len(hits)} möjliga PnP-appar i {t.get('name', 'tenanten')}")
    if v in (None, "__none"):
        return t, None
    return t, v


def ask_client_id(t, default=""):
    """Sök fram PnP-appen i tenanten eller klistra in ClientId. -> (t, ClientId). Cancel om man avbryter."""
    how = choose([dict(label="Sök fram den i tenanten", hint="letar bland appregistreringarna via Graph", value="find"),
                  dict(label="Klistra in ClientId", hint="Application ID, en GUID", value="paste")],
                 title="Var finns PnP-appen?")
    if how is None:
        raise Cancel()
    if how == "find":
        t2, cid = find_pnp_app(t)
        t = dict(t2, client_id=t.get("client_id", "")) if t2 else t
        if cid:
            out("    " + green("→ ") + f"ClientId {cid}")
            return t, cid
        out(dim("  Då klistrar vi in den i stället."))
    cid = ask_text("ClientId (Application ID, en GUID)", default, required=True,
                   check=lambda v: None if GUID.match(v) else "Det ser inte ut som en GUID.")
    return t, cid


def import_cert(t):
    path = ask_text("Sökväg till .pfx-filen", required=True,
                    check=lambda v: None if Path(os.path.expanduser(v.strip('"'))).is_file() else "Hittar ingen fil där.")
    pw = ask_text("Lösenordet till .pfx-filen (syns inte, sparas i valvet)", secret=True)
    CERTS.mkdir(parents=True, exist_ok=True)
    private(CERTS)
    dst = CERTS / (slug(t["name"]) + ".pfx")
    shutil.copyfile(os.path.expanduser(path.strip('"')), dst)
    private(dst)
    if not sec_set(t, pw):
        say("Kunde inte spara lösenordet i valvet.", "sad")
        return False
    t["cert_path"] = str(dst)
    out("  " + green("✔") + f" Certifikatet kopierat till {dst} · lösenordet i valvet {VAULT}")
    return True


def register_app(t):
    app = ask_text("Appens namn i Entra ID", f"PnP PowerShell ({t['name']})")
    perm = "read"
    if t["auth"] == "certificate":
        perm = choose([dict(label=v[0], hint=v[1], value=k) for k, v in PERMS.items()], title="Rättigheter för appen")
        if perm is None:
            raise Cancel()
    outdir = CERTS / ("ny-" + slug(t["name"]))
    code = register_code(app, t["tenant"], t["auth"], perm, outdir)
    print()
    show_command(code, "Registrerar app")
    say(f"Det här skapar en app-registrering i {t['tenant']} och kräver en admin som godkänner (samtycke). "
        "Ett webbläsarfönster öppnas.", "worry")
    if DRY or not confirm(f"Registrera appen i {t['tenant']} på riktigt?", False):
        say("Ingen app skapad. Du kan köra kommandot själv och sedan välja 'Det finns redan en app'.", "happy")
        return False
    secret = secrets.token_urlsafe(24) if t["auth"] == "certificate" else None
    if secret:
        if not store_ready():
            return False
        CERTS.mkdir(parents=True, exist_ok=True)
        private(CERTS)
        outdir.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix="puppa-", suffix=".json")
    os.close(fd)
    env = dict(os.environ, PUPPA_OUT=tmp)
    if secret:
        env["PUPPA_SECRET"] = secret
    out(dim("  Lämnar över till PowerShell. Följ instruktionerna där…"))
    rc = subprocess.call(ps_args(REG_WRAP.format(cmd=code)), env=env)
    data = load_json(tmp, {})
    os.unlink(tmp)
    cid = next((v for k, v in data.items() if re.search(r"clientid|appid", k, re.I) and GUID.match(str(v))), "")
    if rc != 0 or not cid:
        say("Registreringen gick inte igenom (eller så hittade jag inget ClientId).", "sad")
        return False
    t["client_id"] = cid
    if secret:
        pfx = next(iter(outdir.glob("*.pfx")), None)
        if not pfx:
            say("Appen skapades men jag hittar ingen .pfx-fil.", "sad")
            return False
        dst = CERTS / (slug(t["name"]) + ".pfx")
        shutil.move(str(pfx), dst)
        private(dst)
        for f in outdir.glob("*.cer"):
            shutil.move(str(f), CERTS / (slug(t["name"]) + ".cer"))
        shutil.rmtree(outdir, ignore_errors=True)
        if not sec_set(t, secret):
            say("Kunde inte spara certifikatlösenordet i valvet.", "sad")
            return False
        t["cert_path"] = str(dst)
    out("  " + green("✔") + f" App registrerad · ClientId {cid}")
    return True


def print_tenant(t):
    w = width()
    rows = [("Domän", t.get("tenant") or "–"), ("Tenant-ID", t.get("tenant_id") or dim("–")),
            ("Konto", t.get("upn") or dim("–")),
            ("SharePoint", t.get("admin_url") or dim("inte uppsatt")),
            ("PnP-app", t.get("client_id") or dim("ingen – behövs bara för SharePoint-recept"))]
    if t.get("client_id"):
        rows.append(("PnP-inloggning", AUTHS.get(t.get("auth"), "?")))
    if t.get("auth") == "certificate" and t.get("cert_path"):
        rows.append(("Certifikat", t["cert_path"] + dim(f"  · lösenord i valvet {VAULT}")))
    kw = 16 if w >= 58 else 0
    for k, v in rows:
        if kw:
            out("    " + dim(pad(k, kw)) + v)
        else:
            out("    " + dim(k))
            out("      " + v)


def use_tenant(t, quiet=False):
    t = dict(t)
    t["last"] = datetime.now().isoformat(timespec="seconds")
    update_tenant(t["name"], last=t["last"])
    if config().get("active") != t["name"]:
        set_config(active=t["name"])
        if WORKER.conn:
            WORKER.stop()
            WORKER.warm()
    if not quiet:
        say(f"Nu jobbar vi i {t['name']} ({t.get('tenant', '')}).", "proud")


def test_pnp(t):
    print()
    if not t.get("client_id"):
        say(f"{t['name']} har ingen PnP-app än, så det finns inget att testa.", "sad")
        return False
    if config().get("active") != t["name"]:
        use_tenant(t, quiet=True)
    code = connect_code("pnp", t, t["admin_url"]) + "\n" + \
        "$c = Get-PnPConnection\n[pscustomobject]@{ Url = $c.Url; Typ = \"$($c.ConnectionType)\"; " \
        "ClientId = $c.ClientId; Tenant = (Get-PnPTenantId) }"
    show_command(code, "Testar")
    if DRY:
        return True
    try:
        secret = need_secret("pnp", t)
    except RuntimeError as e:
        say(str(e), "sad")
        return False
    ok, n, _ = run_ps(f"Testar {t['name']}", code, secret)
    if ok:
        WORKER.conn["pnp"] = (t["name"], t["admin_url"])
        celebrate(f"{t['name']} fungerar! Du är ansluten till SharePoint-admin.")
    else:
        say("Anslutningen misslyckades. Kolla ClientId, domän och att appen har samtycke.", "sad")
    return ok


def delete_tenant(t):
    say(f"Ta bort {t['name']} ur listan?" + (" Certifikatet och lösenordet i valvet försvinner också."
                                             if t.get("cert_path") else "") + " Inget ändras i själva tenanten.", "worry")
    if not confirm(f"Ta bort {t['name']}?", False):
        return False
    save_tenants([x for x in tenants() if x["name"] != t["name"]])
    if t.get("cert_path"):
        sec_del(t)
    for ext in (".pfx", ".cer"):
        f = CERTS / (slug(t["name"]) + ext)
        if f.exists():
            f.unlink()
    if config().get("active") == t["name"]:
        set_config(active=None)
        WORKER.stop()
        WORKER.warm()
    say(f"{t['name']} är borttagen ur listan.", "happy")
    return True


def set_client_id(t):
    try:
        t2, cid = ask_client_id(t, t.get("client_id", ""))
        t = dict(t2 or t, client_id=cid)
        if not t.get("admin_url"):
            pre = ask_text("SharePoint-namnet? (det som står före .sharepoint.com)", t.get("prefix", ""), required=True)
            t.update(sp_fields(pre.lower()))
        if t.get("auth") == "certificate" and not (t.get("cert_path") and sec_get(t)):
            if not import_cert(t):
                return
    except Cancel:
        out(dim("  Avbrutet."))
        return
    save_tenants([x for x in tenants() if x["name"] != t["name"]] + [t])
    say("ClientId sparat.", "proud")


def open_shell(t):
    """Lämnar över terminalen till ett pwsh-skal som redan är anslutet till tenantens SharePoint-admin."""
    if not t or not t.get("client_id"):
        say("PowerShell-skalet ansluter med PnP, och den här tenanten har ingen PnP-app än.", "worry")
        return
    code = connect_code("pnp", t, t["admin_url"])
    print()
    show_command(code, "Öppnar skal")
    if DRY:
        return
    try:
        secret = need_secret("pnp", t)
    except RuntimeError as e:
        say(str(e), "sad")
        return
    env = dict(os.environ)
    if secret:
        env["PUPPA_SECRET"] = secret
    script = code + "\nRemove-Item Env:PUPPA_SECRET -ErrorAction SilentlyContinue\n" + \
        f"Write-Host 'Ansluten till {t['name'].replace(chr(39), chr(39) * 2)} – skriv exit för att gå tillbaka till Puppa.' -ForegroundColor Cyan"
    say(f"Öppnar PowerShell anslutet till {t['name']}. Skriv exit för att komma tillbaka.", "happy")
    subprocess.call(ps_args(script, "-NoExit"), env=env)


def tenants_menu():
    """Listan över tenants du loggat in i. Välj en för att jobba där, eller hantera den."""
    while True:
        ts = tenants()
        rows = tenant_rows()
        if rows:
            rows.append(("head", ""))
        rows += [dict(label="Lägg till tenant…", hint="oftast inte nödvändigt – inloggning sparar automatiskt",
                      value="__new", key="n"),
                 dict(label="Ny PnP-appregistrering…", hint="för SharePoint-recept", value="__pnp", key="p"),
                 dict(label="Tillbaka", value=None, key="q")]
        name = choose(rows, title="Tenants" + (f" · {len(ts)} st" if ts else ""), status=conn_status, screen=True,
                      msg="" if ts else "Listan fylls på av sig själv när du loggar in någonstans. "
                                        "Kör ett recept, eller lägg till en tenant här.",
                      footer="enter väljer · n ny · p PnP-app · q tillbaka")
        if name is None:
            return
        try:
            if name == "__new":
                add_tenant()
                continue
            if name == "__pnp":
                pnp_setup(active_tenant())
                continue
        except Cancel:
            out(dim("  Avbrutet."))
            continue
        t = find_tenant(name)
        if not t:
            continue
        out(f"  {accent('▍')}{bold(t['name'])}")
        print_tenant(t)
        has = bool(t.get("client_id"))
        acts = [dict(label="Jobba här", hint="gör den aktiv", value="use")]
        if has:
            acts += [dict(label="Testa PnP-anslutningen", hint="Connect-PnPOnline mot SharePoint-admin", value="test"),
                     dict(label="PowerShell-skal (PnP)", hint="anslutet, du tar över", value="shell"),
                     dict(label="Byt ClientId…", value="cid")]
        else:
            acts += [dict(label="Ny PnP-appregistrering…", hint="behövs för SharePoint-recept", value="pnp"),
                     dict(label="Hitta befintlig PnP-app…", hint="sök i tenanten eller ange ClientId", value="cid")]
        acts += [dict(label="Ta bort ur listan…", value="del"), dict(label="Tillbaka", value=None, key="q")]
        act = choose(acts, title=t["name"])
        if act == "use":
            use_tenant(t, quiet=True)
            return
        if act == "test":
            test_pnp(t)
        elif act == "shell":
            open_shell(t)
        elif act == "pnp":
            pnp_setup(t)
        elif act == "cid":
            set_client_id(t)
        elif act == "del":
            delete_tenant(t)


LOGOUT_PARTS = {
    "pnp": ("SharePoint (PnP)", "Disconnect-PnPOnline -ClearPersistedLogin",
            "Import-Module PnP.PowerShell -ErrorAction Stop; Disconnect-PnPOnline -ClearPersistedLogin -ErrorAction Stop"),
    "exo": ("Exchange Online", "Disconnect-ExchangeOnline -Confirm:$false",
            "Import-Module ExchangeOnlineManagement -ErrorAction Stop; Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop *> $null"),
    "teams": ("Teams", "Disconnect-MicrosoftTeams -Confirm:$false",
              "Import-Module MicrosoftTeams -ErrorAction Stop; Disconnect-MicrosoftTeams -Confirm:$false -ErrorAction Stop *> $null"),
    "graph": ("Microsoft Graph", "Disconnect-MgGraph",
              "Import-Module Microsoft.Graph.Authentication -ErrorAction Stop; Disconnect-MgGraph -ErrorAction Stop *> $null"),
}


def logout(parts=None):
    """Logga ut från valda tjänster (standard: alla) och rensa deras sparade inloggningar."""
    parts = [x for x in (parts or LOGOUT_PARTS) if x in LOGOUT_PARTS]
    print()
    show_command("\n".join(LOGOUT_PARTS[x][1] for x in parts), "Loggar ut")
    if DRY:
        return True
    title = "Loggar ut från " + (", ".join(LOGOUT_PARTS[x][0] for x in parts)
                                 if len(parts) < len(LOGOUT_PARTS) else "allt")
    lines = []
    # PnP och övriga tjänster bor i var sin PowerShell-process (se Workers) – en körning per process
    for group in ([x for x in parts if x == "pnp"], [x for x in parts if x != "pnp"]):
        if group:
            code = "\n".join(f"try {{ {LOGOUT_PARTS[x][2]}; Write-Host 'utloggad: {LOGOUT_PARTS[x][0]}' }} catch {{ }}"
                             for x in group)
            lines += run_ps(title, code, show_output=False)[2]
    done = [t.split("utloggad: ", 1)[1] for k, t in lines if k == "H" and "utloggad: " in t]
    if len(parts) == len(LOGOUT_PARTS):
        WORKER.stop()
        WORKER.warm()
    for x in parts:
        WORKER.conn.pop(x, None)
    LOOKUP_CACHE.clear()
    for d in done:
        out("  " + green("✔") + " " + d)
    say(("Utloggad från allt." if len(parts) == len(LOGOUT_PARTS) else "Utloggad.") +
        " Nästa gång frågar jag efter inloggning igen. (Webbläsarens Microsoft-inloggning ligger kvar – "
        "logga ut där också om datorn delas.)", "happy")
    return True


def logout_menu(at_exit=False):
    """Välj vad du loggar ut från. Allt är förvalt."""
    t = active_tenant()
    rows = [dict(label="Allt", hint="SharePoint · Exchange · Teams · Graph", value="all")]
    for k, (name, cmd, _) in LOGOUT_PARTS.items():
        on = bool(t) and WORKER.conn.get(k, ("",))[0] == t["name"]
        rows.append(dict(label=f"Bara {name}", hint=("inloggad nu · " if on else "") + cmd, value=k))
    rows += [("head", ""), dict(label="Nej, förbli inloggad" if at_exit else "Tillbaka", value="__no", key="n" if at_exit else "q")]
    v = choose(rows, title="Logga ut innan du går?" if at_exit else "Logga ut från…", status=None if at_exit else conn_status,
               screen=not at_exit)
    if v in (None, "__no"):
        return False
    return logout(None if v == "all" else [v])


def connect_menu():
    t = active_tenant()
    who = t["name"] if t else "välj konto vid inloggning"
    rows = [dict(label="Exchange Online", hint="Connect-ExchangeOnline", value="exo"),
            dict(label="Microsoft Graph…", hint="Connect-MgGraph med valda scopes", value="graph"),
            dict(label="Teams", hint="Connect-MicrosoftTeams", value="teams")]
    if t and t.get("client_id"):
        rows += [dict(label="SharePoint-admin (PnP)", hint=t["admin_url"], value="pnp-admin"),
                 dict(label="SharePoint-sajt (PnP)…", hint="valfri sajt-URL", value="pnp-site")]
    else:
        rows.append(dict(label="SharePoint (PnP)…", hint="kräver en PnP-appregistrering i tenanten", value="pnp-admin"))
    keep = keep_logins()
    rows += [dict(label="Logga ut…", hint="allt eller en tjänst i taget", value="off"),
             dict(label="Kom ihåg inloggningar: " + ("på" if keep else "av"),
                  hint=("sparas i nyckelringen mellan körningar" if keep else "bara medan Puppa är igång – ingen nyckelringsfråga"),
                  value="keep"),
             dict(label="Tillbaka", value=None, key="q")]
    v = choose(rows, title=f"Logga in · {who}", status=conn_status, screen=True)
    if v == "keep":
        set_config(keep_logins=not keep)
        if not keep:
            say("Inloggningar sparas nu mellan körningar (Graph och PnP lägger tokens i nyckelringen). "
                "Får du en nyckelringsfråga varje gång: välj 'Tillåt alltid', eller installera en signerad pwsh "
                "(brew install --cask powershell).", "happy")
        else:
            say("Inloggningar sparas inte längre – du loggar in en gång per Puppa-session och slipper nyckelringen.", "happy")
        return
    try:
        if v == "pnp-admin":
            ensure("pnp", t)
        elif v == "pnp-site":
            url = ask_text("Sajtens URL", t["root_url"] + "/sites/", required=True)
            ensure("pnp", t, url)
        elif v in ("exo", "teams"):
            ensure(v, t)
        elif v == "graph":
            sc = ask_text("Scopes (kommaseparerade)", "User.Read.All,Group.Read.All,Organization.Read.All")
            ensure("graph", t, scopes=[s.strip() for s in sc.split(",") if s.strip()])
        elif v == "off":
            logout_menu()
    except Cancel:
        out(dim("  Avbrutet."))


CAT_HINTS = {"SharePoint": "sajter, behörigheter, papperskorg, delning", "OneDrive": "hitta, ge åtkomst, största",
             "Exchange": "delade postlådor, åtkomst, kalender, listor, spårning",
             "Grupper": "medlemmar, lägg till/ta bort, utan ägare", "Användare": "info, licenser, gäster, MFA, spärra",
             "Teams": "team, medlemmar, kanaler", "Entra ID": "inloggningar, granskning, villkorlig åtkomst, enheter, appar, roller",
             "Excel": "exportera senaste resultatet", "Lokalt": "moduler, session"}


def remember_recipe(rid):
    rec = [x for x in config().get("recent", []) if x != rid]
    set_config(recent=([rid] + rec)[:5])


def recipe_rows(rs):
    rows = [dict(label=r.name, hint=r.desc, value=r.id, badge=r.badge()) for r in rs if not r.writes]
    w = [dict(label=r.name, hint=r.desc, value=r.id, badge=r.badge()) for r in rs if r.writes]
    if rows and w:
        rows.append(("head", "ändrar"))
    return rows + w


def recipe_search_rows(rs=None):
    """Alla recept som sökbara rader (värde 'r:<id>') – för fuzzy-sökning i menyerna."""
    return [dict(label=r.name, hint=f"{r.category} · {r.desc}", value="r:" + r.id, badge=r.badge(),
                 search=r.id) for r in (rs or load_recipes())]


def run_recipe_by_value(v):
    r = find_recipe(v[2:])
    if r:
        remember_recipe(r.id)
        run_recipe(r)


def recipes_menu():
    """Två steg: kategori → recept. Efter körning stannar man i samma kategori."""
    cat_last = None
    while True:
        rs = load_recipes()
        cats = []
        for r in rs:
            if r.category not in cats:
                cats.append(r.category)
        by_id = {r.id: r for r in rs}
        recent = [by_id[x] for x in config().get("recent", []) if x in by_id]
        rows = []
        if recent:
            rows.append(dict(label="Senast använda", hint=", ".join(r.name for r in recent[:3]), value="__recent",
                             badge=dim(str(len(recent)))))
        for cname in cats:
            its = [r for r in rs if r.category == cname]
            rows.append(dict(label=cname, hint=CAT_HINTS.get(cname, ""), value=cname, badge=dim(str(len(its)))))
        rows += [("head", ""), dict(label="Nytt recept (mall)…", hint="skriv ett eget", value="__new", key="n"),
                 dict(label="Tillbaka", value=None, key="q")]
        cat = choose(rows, title="Recept · välj område", start=cat_last, status=conn_status, screen=True,
                     search_rows=recipe_search_rows(rs))
        if cat is None:
            return
        if cat.startswith("r:"):
            run_recipe_by_value(cat)
            continue
        if cat == "__new":
            new_recipe()
            continue
        cat_last = cat
        its = recent if cat == "__recent" else [r for r in rs if r.category == cat]
        last = None
        while True:
            title = "Senast använda" if cat == "__recent" else cat
            v = choose(recipe_rows(its) + [("head", ""), dict(label="Tillbaka", value=None, key="q")],
                       title=title, start=last, status=conn_status, screen=True,
                       footer="↑/↓ · enter kör · siffra väljer · q tillbaka" if width() >= 50 else None)
            if not v:
                break
            last = v
            remember_recipe(v)
            run_recipe(by_id[v])


# ───────────────────────────── huvudmeny ─────────────────────────────
def greeting():
    ts = tenants()
    t = active_tenant()
    if not ts:
        return ("Hej! Jag är Puppa. Välj ett recept – jag loggar in när det behövs och lägger tenanten i "
                "listan, så kan du växla sen."), "happy"
    if not t:
        return "Välj en tenant i listan, eller kör ett recept så loggar jag in.", "happy"
    if any(WORKER.conn.get(k, ("",))[0] == t["name"] for k in CONN_NAMES):
        return "Inloggad och redo. Vad blir nästa steg?", "happy"
    return "Välkommen tillbaka! Fortsätt där du lämnade, eller börja skriva för att söka bland recepten.", "happy"


def main_menu():
    last = None
    while True:
        msg, mood = greeting()
        ts = tenants()
        t = active_tenant()
        rows = []
        by_id = {r.id: r for r in load_recipes()}
        recent = [by_id[x] for x in config().get("recent", []) if x in by_id][:3]
        if t and recent:
            rows.append(("head", f"Fortsätt där du lämnade · {t['name']}"))
            rows += [dict(label=r.name, hint=r.category, value="r:" + r.id, badge=r.badge()) for r in recent]
            rows.append(("head", ""))
        rows += [dict(label="Tenants…", hint=(f"{t['name']} · " if t else "") +
                     (f"{len(ts)} sparade – välj vilken du jobbar i" if ts else "listan fylls på när du loggar in"),
                     value="tenants", key="t"),
                dict(label="Recept…", hint="kör en uppgift: lista, skapa, exportera", value="recipes"),
                dict(label="Logga in…", hint="Exchange · Graph · Teams · SharePoint (PnP)", value="connect"),
                dict(label="Ny PnP-appregistrering", hint="behövs bara för SharePoint-recept", value="pnp")]
        if t and t.get("client_id"):
            rows.append(dict(label="PowerShell-skal (PnP)", hint="anslutet till tenanten, du tar över", value="shell"))
        rows += [dict(label="Logga ut…", hint="allt eller en tjänst i taget", value="logout", key="u"),
                 dict(label="Avsluta", value="quit", key="q")]
        v = choose(rows, msg=msg, mood=mood, status=conn_status,
                   start=last or ("r:" + recent[0].id if t and recent else "recipes"),
                   footer="↑/↓ · enter · siffra/bokstav väljer · q avslutar",
                   sub=f"puppa {VERSION}" + ("  ·  torrkörning" if DRY else ""), screen=True,
                   search_rows=recipe_search_rows())
        last = v
        if v in (None, "quit"):
            print()
            if WORKER.conn:
                try:
                    logout_menu(at_exit=True)
                except Cancel:
                    pass
            say(random.choice(["Hej då!", "Vi ses!", "Hej då! -WhatIf var med dig."]), "happy")
            print()
            return
        try:
            if v.startswith("r:"):
                run_recipe_by_value(v)
            elif v == "recipes":
                recipes_menu()
            elif v == "tenants":
                tenants_menu()
            elif v == "pnp":
                pnp_setup(t)
            elif v == "logout":
                logout_menu()
            elif v == "connect":
                connect_menu()
            elif v == "shell":
                open_shell(t)
        except Cancel:
            out(dim("  Avbrutet."))


# ───────────────────────────── CLI ─────────────────────────────
def list_tenants():
    ts = tenants()
    if not ts:
        say("Inga tenants än. De sparas automatiskt när du loggar in via ett recept (eller: puppa ny).", "happy")
        return
    act = config().get("active")
    w = width()
    nw = min(max(len(t["name"]) for t in ts) + 2, 24)
    for t in sorted(ts, key=lambda x: x.get("last") or x.get("created") or "", reverse=True):
        mark = green("●") if t["name"] == act else dim("○")
        pnp = cyan("PnP") if t.get("client_id") else dim("–")
        if w >= 72:
            out(f"  {mark} {pad(bold(cut(t['name'], nw)), nw)}{pad(cut(t.get('tenant', ''), 32), 34)}{pad(pnp, 5)}"
                + dim(cut(t.get("upn", ""), max(4, w - nw - 46))))
        elif w >= 44:
            out(f"  {mark} {pad(bold(cut(t['name'], nw)), nw)}{dim(cut(t.get('tenant', ''), w - nw - 6))}")
        else:
            out(f"  {mark} {bold(t['name'])}")
            out(f"    {dim(cut(t.get('tenant', ''), w - 6))}")


def list_recipes():
    rs = load_recipes()
    w = width()
    nw = min(max(len(r.id) for r in rs) + 2, 36)
    cat = None
    for r in rs:
        if r.category != cat:
            cat = r.category
            out("  " + accent("▍") + bold(cat))
        if w >= 58:
            out(f"    {pad(cut(r.id, nw), nw)}{pad(r.badge(), 8)}" + (dim(r.name) if w >= 80 else ""))
        else:
            out(f"    {cut(r.id, w - 14)} {r.badge()}")


def main(argv=None):
    import argparse
    global ANIM, DRY
    argv = list(sys.argv[1:] if argv is None else argv)
    pre = argparse.ArgumentParser(add_help=False)
    pre.add_argument("--no-anim", action="store_true")
    pre.add_argument("-n", "--torr", action="store_true")
    known, rest = pre.parse_known_args(argv)
    if known.no_anim:
        ANIM = False
    DRY = known.torr

    ap = argparse.ArgumentParser(prog="puppa", description="Puppa – M365-admin i PowerShell med några knapptryck.")
    ap.add_argument("--version", action="version", version="puppa " + VERSION)
    sub = ap.add_subparsers(dest="cmd")
    sub.add_parser("tenants", aliases=["kunder", "profiler"], help="tenants du loggat in i")
    u = sub.add_parser("tenant", aliases=["kund", "använd", "anvand", "use"], help="växla tenant")
    u.add_argument("namn")
    sub.add_parser("ny", help="lägg till en tenant för hand")
    pp = sub.add_parser("pnp", help="ny PnP-appregistrering för en tenant")
    pp.add_argument("namn", nargs="?")
    t = sub.add_parser("test", help="testa PnP-anslutning")
    t.add_argument("namn", nargs="?")
    s = sub.add_parser("skal", aliases=["shell"], help="öppna anslutet pwsh-skal (PnP)")
    s.add_argument("namn", nargs="?")
    sub.add_parser("recept", help="lista recept")
    k = sub.add_parser("kör", aliases=["kor", "run"], help="kör ett recept")
    k.add_argument("recept")
    k.add_argument("par", nargs="*", help="Namn=värde")
    k.add_argument("--ja", action="store_true", help="bekräfta ändrande recept utan att fråga")
    sub.add_parser("logga-ut", aliases=["logout"], help="logga ut från PnP, Exchange, Teams och Graph")
    sub.add_parser("hej")
    a = ap.parse_args(rest)

    if not (shutil.which("pwsh") or shutil.which("pwsh-preview")):
        say("Jag hittar inte pwsh (PowerShell 7). Installera det: Windows: winget install Microsoft.PowerShell · "
            "macOS: signerat .pkg från github.com/PowerShell/PowerShell/releases (Homebrew-formeln är osignerad "
            "och ger nyckelringsfrågor) · eller: uppa install powershell", "sad")
        return 1

    def pick(name):
        t = find_tenant(name) if name else active_tenant()
        if not t:
            say(f"Hittar ingen tenant {name or '(ingen vald)'}. Se puppa tenants.", "sad")
            sys.exit(1)
        return t

    try:
        if a.cmd is None:
            if not ITTY or not TTY:
                print(__doc__)
                return 0
            signal.signal(signal.SIGINT, signal.default_int_handler)
            WORKER.warm()
            main_menu()
            return 0
        if a.cmd in ("tenants", "kunder", "profiler"):
            list_tenants()
            return 0
        if a.cmd in ("tenant", "kund", "använd", "anvand", "use"):
            use_tenant(pick(a.namn))
            return 0
        if a.cmd == "ny":
            if not ITTY:
                sys.exit("  behöver en riktig terminal")
            return 0 if add_tenant() else 1
        if a.cmd == "pnp":
            if not ITTY:
                sys.exit("  guiden behöver en riktig terminal")
            WORKER.warm()
            return 0 if pnp_setup(find_tenant(a.namn) if a.namn else active_tenant()) else 1
        if a.cmd == "test":
            WORKER.warm()
            return 0 if test_pnp(pick(a.namn)) else 1
        if a.cmd in ("skal", "shell"):
            open_shell(pick(a.namn))
            return 0
        if a.cmd == "recept":
            list_recipes()
            return 0
        if a.cmd in ("kör", "kor", "run"):
            r = find_recipe(a.recept)
            if not r:
                hits = [x.id for x in load_recipes() if a.recept.lower() in x.id.lower() or a.recept.lower() in x.name.lower()]
                say(f"Flera recept matchar {a.recept}: {', '.join(hits[:6])}." if hits else
                    f"Hittar inget recept som heter {a.recept}. Se puppa recept.", "worry" if hits else "sad")
                return 1
            given = dict(x.split("=", 1) for x in a.par if "=" in x)
            WORKER.warm()
            return 0 if run_recipe(r, given, yes=a.ja) else 1
        if a.cmd in ("logga-ut", "logout"):
            return 0 if logout() else 1
        if a.cmd == "hej":
            print()
            say(random.choice(JOKES), random.choice(["happy", "proud", "wow", "left"]))
            print()
            return 0
    except Cancel:
        out(dim("  Avbrutet."))
        return 1
    except KeyboardInterrupt:
        sys.stdout.write("\033[?25h\n")
        return 130
    finally:
        WORKER.stop()
    return 0


if __name__ == "__main__":
    sys.exit(main())

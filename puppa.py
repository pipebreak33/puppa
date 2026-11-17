#!/usr/bin/env python3
"""puppa – vanliga M365-adminuppgifter i PowerShell med några knapptryck.

  puppa                    meny
  puppa ny                 guide: ny PnP-profil (registrera Entra ID-app eller ange ClientId)
  puppa profiler           lista sparade profiler
  puppa använd PROFIL      gör en profil aktiv
  puppa test [PROFIL]      testa PnP-anslutningen
  puppa skal [PROFIL]      öppna ett pwsh-skal som redan är anslutet med Connect-PnPOnline
  puppa recept             lista recept
  puppa kör RECEPT [Namn=värde …] [--ja]   kör ett recept (frågar efter det som saknas)
  puppa logga-ut           logga ut från PnP, Exchange, Teams och Graph (rensar sparade inloggningar)
  puppa hej

Flaggor: -n/--torr visar kommandon utan att köra något · --no-anim stänger av animationer.
Profiler: ~/.puppa/profiles.json (inga hemligheter). Certifikatlösenord: macOS Keychain (tjänst "puppa").
Egna recept: ~/.puppa/recipes/*.ps1. Enbart standardbiblioteket, Python 3.8+, kräver pwsh.
"""
from __future__ import annotations

import json
import os
import queue
import random
import re
import secrets
import select
import shutil
import signal
import subprocess
import sys
import tempfile
import threading
import time
from datetime import datetime
from pathlib import Path

VERSION = "1.0.0"
HOME = Path.home()
DIR = Path(os.environ.get("PUPPA_DIR") or HOME / ".puppa")
os.environ["PUPPA_DIR"] = str(DIR)
PROFILES = DIR / "profiles.json"
CONFIG = DIR / "config.json"
CERTS = DIR / "certs"
USER_RECIPES = DIR / "recipes"
APP_DIR = Path(os.path.realpath(__file__)).parent
RECIPES = APP_DIR / "recipes"
EXPORTS = HOME / "Documents" / "Puppa"
KC_SERVICE = "puppa"
for _p in ("/opt/homebrew/bin", "/usr/local/bin", "/usr/local/microsoft/powershell/7"):
    if _p not in os.environ.get("PATH", "").split(os.pathsep):
        os.environ["PATH"] = os.environ.get("PATH", "") + os.pathsep + _p

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


def raw_keys():
    import termios
    import tty
    fd = sys.stdin.fileno()
    old = termios.tcgetattr(fd)
    tty.setcbreak(fd)
    return fd, old


def restore_keys(fd, old):
    import termios
    termios.tcsetattr(fd, termios.TCSADRAIN, old)


def new_screen():
    """Töm skärmen inför en meny. Har något just skrivits ut får man läsa klart först."""
    global DIRTY
    if DIRTY:
        pause()
    DIRTY = False
    if TTY:
        sys.stdout.write("\033[H\033[2J")
        sys.stdout.flush()


def choose(rows, title="", msg="", mood="happy", sub="", status=None, start=None, footer=None, screen=False):
    """Piltangent-meny. rows: ("head", text) eller dict(label, hint, value, badge, key).
    Siffra väljer direkt, q/esc avbryter (→ None). Maskoten blinkar medan den väntar.
    screen=True: menyn ritas på en tom skärm (navigeringsmenyer), annars där markören står (frågor i ett flöde)."""
    if not ITTY:
        raise Cancel()
    if screen:
        new_screen()
    sel = [k for k, r in enumerate(rows) if not isinstance(r, tuple)]
    if not sel:
        return None
    keys = {}
    n = 0
    for k in sel:
        r = rows[k]
        if r.get("key"):
            keys[r["key"]] = k
        else:
            n += 1
            if n <= 9:
                r["_num"] = str(n)
                keys[str(n)] = k
    pos = sel[0]
    if start is not None:
        for k in sel:
            if rows[k]["value"] == start:
                pos = k
    top = 0
    live = Live()
    winch = {"flag": False}
    t0 = time.time()
    state = {"mood": None}

    def draw(mood_now):
        nonlocal top
        w = width()
        lines = []
        if msg:
            lines += frame(msg, mood_now, sub=sub)
        if status:
            lines += [""] + status()
        lines.append("")
        if title:
            lines.append("  " + accent("▍") + bold(title))
        narrow = w < 58
        labels = [rows[k]["label"] for k in sel]
        lw = min(max(len(x) for x in labels) + 2, max(10, w - 14 if narrow else min(34, w // 2)))
        room = height() - len(lines) - 3
        room = max(3, room)
        top = min(max(top, pos - room + 1), pos)
        top = max(0, min(top, max(0, len(rows) - room)))
        if top > 0:
            lines.append(dim("     ↑ fler"))
        for k in range(top, min(len(rows), top + room)):
            r = rows[k]
            if isinstance(r, tuple):
                lines.append("   " + dim(cut(r[1].upper(), w - 5)))
                continue
            on = k == pos
            key = r.get("key") or r.get("_num") or " "
            label = cut(r["label"], lw - 1)
            label = bold(accent(label)) if on else label
            badge = r.get("badge") or ""
            if narrow:
                line = f"  {accent('❯') if on else ' '}{dim(key)} {label}" + (" " + badge if badge else "")
            else:
                line = f"   {accent('❯') if on else ' '} {dim(key)}  " + pad(label, lw) + \
                       (pad(badge, 8) if badge or any(rows[j].get("badge") for j in sel) else "")
                if r.get("hint") and w >= 64:
                    line += dim(r["hint"])
            lines.append(line)
        if top + room < len(rows):
            lines.append(dim("     ↓ fler"))
        lines.append("")
        help_ = (footer or "↑/↓ · enter · siffra väljer · q tillbaka") if w >= 50 else "↑↓ enter · q"
        lines.append(dim("   " + help_))
        live.draw(lines)

    def on_winch(*_):
        winch["flag"] = True

    old_winch = signal.signal(signal.SIGWINCH, on_winch)
    fd, old = raw_keys()
    choice = None
    try:
        sys.stdout.write("\033[?25l")
        draw(mood)
        while True:
            r, _, _ = select.select([fd], [], [], 0.12)
            if not r:
                m = mood
                if ANIM:
                    im = idle_mood(time.time() - t0)
                    m = im if mood == "happy" else ("blink" if im == "blink" else mood)
                if winch["flag"] or m != state["mood"]:
                    winch["flag"] = False
                    state["mood"] = m
                    draw(m)
                continue
            ch = os.read(fd, 8).decode(errors="ignore")
            idx = sel.index(pos)
            if ch in ("\x1b[A", "k", "\x1bOA"):
                pos = sel[(idx - 1) % len(sel)]
            elif ch in ("\x1b[B", "j", "\x1bOB"):
                pos = sel[(idx + 1) % len(sel)]
            elif ch in keys:
                pos = keys[ch]
                choice = rows[pos]["value"]
                break
            elif ch in ("\r", "\n"):
                choice = rows[pos]["value"]
                break
            elif ch in ("q", "\x1b", "\x03", "\x04"):
                break
            draw(state["mood"] or mood)
    except KeyboardInterrupt:
        choice = None
    finally:
        signal.signal(signal.SIGWINCH, old_winch)
        restore_keys(fd, old)
        live.clear()
        sys.stdout.write("\033[?25h")
        sys.stdout.flush()
    if choice is not None and not screen:
        out(f"   {accent('❯')} {bold(rows[pos]['label'])}")
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
    if ITTY:
        try:
            print()
            input(dim("   enter = tillbaka till menyn "))
        except (EOFError, KeyboardInterrupt):
            print()


# ───────────────────────────── lagring ─────────────────────────────
def load_json(path, default):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def save_json(path, data):
    DIR.mkdir(parents=True, exist_ok=True)
    os.chmod(DIR, 0o700)
    tmp = Path(str(path) + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=1, ensure_ascii=False)
    os.chmod(tmp, 0o600)
    os.replace(tmp, path)


def config():
    return load_json(CONFIG, {})


def set_config(**kw):
    cfg = config()
    cfg.update(kw)
    save_json(CONFIG, cfg)


def profiles():
    return load_json(PROFILES, {"profiles": []}).get("profiles", [])


def save_profiles(ps):
    save_json(PROFILES, {"profiles": ps})


def find_profile(name):
    if not name:
        return None
    for p in profiles():
        if p["name"].lower() == name.lower():
            return p
    hits = [p for p in profiles() if p["name"].lower().startswith(name.lower())]
    return hits[0] if len(hits) == 1 else None


def active_profile():
    return find_profile(config().get("active"))


def slug(s):
    return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-") or "profil"


# macOS Keychain: bara certifikatlösenord. `security -i` så att hemligheten aldrig syns i processlistan.
def kc_q(s):
    return s.replace("\\", "\\\\").replace('"', '\\"')


def kc_set(account, secret):
    cmd = f'add-generic-password -U -s {KC_SERVICE} -a "{kc_q(account)}" -l "Puppa {kc_q(account)}" -w "{kc_q(secret)}"\n'
    subprocess.run(["security", "-i"], input=cmd, text=True, capture_output=True)
    return kc_get(account) == secret


def kc_get(account):
    p = subprocess.run(["security", "find-generic-password", "-s", KC_SERVICE, "-a", account, "-w"],
                       capture_output=True, text=True)
    return p.stdout.rstrip("\n") if p.returncode == 0 else None


def kc_del(account):
    subprocess.run(["security", "delete-generic-password", "-s", KC_SERVICE, "-a", account], capture_output=True)


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
    """En långlivad pwsh-process: anslutningar (PnP, EXO, Teams, Graph) lever kvar mellan recepten."""

    def __init__(self):
        self.p = None
        self.q = queue.Queue()
        self.conn = {}
        self.lock = threading.Lock()
        self.ready = threading.Event()

    def start(self):
        with self.lock:
            if self.p and self.p.poll() is None:
                return
            self.conn = {}
            self.ready.clear()
            DIR.mkdir(parents=True, exist_ok=True)
            os.chmod(DIR, 0o700)
            self.q = queue.Queue()
            self.p = subprocess.Popen(["pwsh", "-NoLogo", "-NoProfile", "-NonInteractive", "-Command", WORKER_PS],
                                      stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                      text=True, encoding="utf-8", errors="replace", bufsize=1,
                                      start_new_session=True)
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
        self.conn = {}

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
            else:
                on_event("H", ln)


WORKER = Worker()


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
            if kind in ("H", "W", "E") and text.strip():
                self.tail = (self.tail + [text])[-3:]
                m = self.DEVICE.search(text)
                if m:
                    code = m.group(2)
                    try:
                        subprocess.run(["pbcopy"], input=code, text=True)
                    except Exception:
                        pass
                    self.msg = f"Logga in på {m.group(1).split('//')[-1]} med koden {code} (kopierad)."
        if not ANIM and kind == "H":
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
    (r"AADSTS700016", "ClientId:t finns inte i den här tenanten. Fel tenant i profilen, eller så är appen inte registrerad där."),
    (r"AADSTS65001|consent", "Appen saknar samtycke. En Global Admin måste godkänna den (Entra ID → Appregistreringar → API-behörigheter)."),
    (r"AADSTS50020|AADSTS50034", "Kontot finns inte i kundens tenant. Logga in med ett konto hos kunden."),
    (r"AADSTS7000215|AADSTS700027|certificate", "Certifikatet godtas inte. Kontrollera att rätt certifikat är uppladdat på appen."),
    (r"\(403\)|Forbidden|Access is denied|Åtkomst nekad", "Behörighet saknas. Kontot/appen har inte rätt roll eller rättighet för det här."),
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
    errs = [t for k, t in lines if k == "E"]
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
        out(dim(f"  … +{len(body) - limit} rader. Kör receptet 'Exportera till Excel' för allt."))
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


def pnp_connect_code(p, url):
    base = f"Connect-PnPOnline -Url {psq(url)} -ClientId {psq(p['client_id'])} -Tenant {psq(p['tenant'])}"
    a = p.get("auth", "interactive")
    if a == "certificate":
        return base + f" -CertificatePath {psq(p['cert_path'])} -CertificatePassword {SECRET_EXPR} -ErrorAction Stop"
    if a == "device":
        return base + " -DeviceLogin -PersistLogin -ErrorAction Stop"
    return base + " -Interactive -PersistLogin -ErrorAction Stop"


def connect_code(kind, p, url=None, scopes=None):
    dev = p.get("auth") == "device"
    if kind == "pnp":
        return pnp_connect_code(p, url or p["admin_url"])
    if kind == "exo":
        s = "Connect-ExchangeOnline -ShowBanner:$false"
        if p.get("upn"):
            s += f" -UserPrincipalName {psq(p['upn'])}"
        if dev:
            s += " -Device"
        return s + " -ErrorAction Stop"
    if kind == "teams":
        s = f"Connect-MicrosoftTeams -TenantId {psq(p['tenant'])}" + (" -UseDeviceAuthentication" if dev else "")
        return s + " -ErrorAction Stop | Out-Null"
    if kind == "graph":
        sc = ",".join(psq(x) for x in sorted(scopes or ["User.Read.All"]))
        return f"Connect-MgGraph -TenantId {psq(p['tenant'])} -Scopes {sc} -NoWelcome" + \
               (" -UseDeviceCode" if dev else "") + " -ErrorAction Stop"
    raise ValueError(kind)


CONN_NAMES = {"pnp": "SharePoint (PnP)", "exo": "Exchange Online", "teams": "Teams", "graph": "Microsoft Graph"}


def need_secret(kind, p):
    if kind == "pnp" and p.get("auth") == "certificate":
        s = kc_get(p["name"])
        if not s:
            raise RuntimeError(f"Hittar inget certifikatlösenord för {p['name']} i Keychain (tjänst 'puppa').")
        return s
    return None


def is_connected(kind, p, url=None, scopes=None):
    cur = WORKER.conn.get(kind)
    if not cur or cur[0] != p["name"]:
        return False
    if kind == "pnp":
        return cur[1].rstrip("/").lower() == (url or p["admin_url"]).rstrip("/").lower()
    if kind == "graph":
        return set(scopes or []) <= set(cur[1])
    return True


def ensure(kind, p, url=None, scopes=None, show=True):
    """Anslut om det behövs. -> True om anslutet."""
    if kind == "pnp" and not p.get("client_id"):
        say(f"Profilen {p['name']} saknar ClientId. Fyll i det under Profiler eller kör guiden.", "sad")
        return False
    if is_connected(kind, p, url, scopes):
        return True
    if kind == "graph" and WORKER.conn.get("graph") and WORKER.conn["graph"][0] == p["name"]:
        scopes = sorted(set(scopes or []) | set(WORKER.conn["graph"][1]))
    code = connect_code(kind, p, url, scopes)
    if show:
        show_command(code, "Ansluter")
    if DRY:
        return True
    try:
        secret = need_secret(kind, p)
    except RuntimeError as e:
        say(str(e), "sad")
        return False
    ok, _, lines = run_ps(f"Ansluter till {CONN_NAMES[kind]} ({p['name']})", code, secret, show_output=False)
    if not ok:
        print_output(lines, quiet_host=not ANIM)
        say(f"Kunde inte ansluta till {CONN_NAMES[kind]}. Felet står ovanför.", "sad")
        return False
    WORKER.conn[kind] = (p["name"], url or p["admin_url"]) if kind == "pnp" else \
        (p["name"], sorted(scopes or [])) if kind == "graph" else (p["name"],)
    out("  " + green("✔") + f" Ansluten till {CONN_NAMES[kind]}" + dim(f" · {p['name']}"))
    return True


def conn_status():
    p = active_profile()
    w = width()
    if not p:
        return ["  " + dim("Ingen aktiv profil")]
    head = f"  {dim('Profil')} {bold(p['name'])}"
    if w >= 58:
        head += dim(f"  {p['tenant']}  ·  {AUTHS.get(p.get('auth'), '?').split(' (')[0].lower()}")
    if not p.get("client_id"):
        head += "  " + yellow("saknar ClientId")
    parts = []
    for k, short in (("pnp", "PnP"), ("exo", "EXO"), ("teams", "Teams"), ("graph", "Graph")):
        on = WORKER.conn.get(k, ("",))[0] == p["name"]
        parts.append((green("● ") if on else dim("○ ")) + (short if on else dim(short)))
    return [head, "  " + "  ".join(parts)]


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
        self.choices = [x.strip() for x in f[3].split(",") if x.strip()] if len(f) > 3 and f[3] else []


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
    order = ["SharePoint", "OneDrive", "Exchange", "Grupper", "Användare", "Teams", "Excel", "Lokalt"]
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


def placeholders(p):
    if not p:
        return {}
    return {"tenant": p.get("tenant", ""), "prefix": p.get("prefix", ""), "adminUrl": p.get("admin_url", ""),
            "rootUrl": p.get("root_url", ""), "upn": p.get("upn", ""), "profil": p.get("name", ""),
            "datum": datetime.now().strftime("%Y-%m-%d")}


def fill(s, ph):
    return re.sub(r"\{(\w+)\}", lambda m: ph.get(m.group(1), m.group(0)), s)


def collect_params(r, p, given):
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
        if prm.choices:
            rows = [dict(label=ch, value=ch) for ch in prm.choices]
            v = choose(rows, title=prm.question, start=default or None)
            if v is None:
                raise Cancel()
        else:
            v = ask_text(prm.question + ("" if not prm.optional else " (valfritt)"), default,
                         required=not prm.optional)
        vals[prm.name] = v
    return vals


def build_script(r, vals, p=None):
    head = "\n".join(f"${k} = {psq(v)}" for k, v in vals.items())
    return (head + "\n" if head else "") + fill(r.body, placeholders(p))


def run_recipe(r, given=None, yes=False):
    p = active_profile()
    needs = r.needs_list()
    if needs and not p:
        say("Det här receptet behöver en kundprofil. Skapa en först – det går fort.", "worry")
        if ITTY and confirm("Starta guiden för ny PnP-profil?", True):
            guide()
        return False
    w = width()
    out(f"  {accent('▍')}{bold(r.name)}  {r.badge()}" + (dim(f"  {p['name']}") if p and needs else ""))
    for l in wrap(r.desc, w - 6, 3) if r.desc else []:
        out("    " + dim(l))
    print()
    try:
        vals = collect_params(r, p, given or {})
    except Cancel as e:
        if e.args:
            out(red("  " + e.args[0]))
        out(dim("  Avbrutet."))
        return False
    script = build_script(r, vals, p)
    print()
    pre = []
    for kind, arg in needs:
        url = None
        scopes = None
        if kind == "pnp":
            url = p["admin_url"] if arg == "admin" else p["root_url"] if arg == "root" else vals.get(arg, arg)
        if kind == "graph":
            scopes = [s.strip() for s in arg.split(",") if s.strip()]
        pre.append((kind, url, scopes))
        if not is_connected(kind, p, url, scopes):
            show_command(connect_code(kind, p, url, scopes) if (kind != "pnp" or p.get("client_id")) else
                         "# profilen saknar ClientId – fyll i under Profiler", "Först: anslut")
    show_command(script)
    if DRY:
        say("Torrkörning: inget kördes. Så här skulle det se ut.", "happy")
        return True
    if r.writes and not yes:
        print()
        say(f"Det här ändrar i {p['tenant'] if p else 'tenanten'}. Kolla kommandot ovan först.", "worry")
        if not confirm("Kör på riktigt?", False):
            say("Okej, inget ändrat.", "happy")
            return False
    for kind, url, scopes in pre:
        if not ensure(kind, p, url, scopes, show=False):
            return False
    ok, n, _ = run_ps(f"Kör {r.name}", script)
    print()
    if ok:
        word = f"{n} rader" if n != 1 else "1 rad"
        tip = " Vill du ha det i Excel? Välj 'Exportera senaste resultat till Excel'." if n > 1 and \
            "excel" not in r.id else ""
        if r.writes:
            celebrate("Klart! Ändringen är gjord.")
        else:
            say(f"Klart! {word}.{tip}" if n else "Klart!", "proud")
    else:
        say("Det gick inte hela vägen. Felet står ovanför i rött.", "sad")
    return ok


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
# Platshållare i standardvärden: {{tenant}} {{prefix}} {{adminUrl}} {{rootUrl}} {{upn}} {{profil}} {{datum}}
# Returnera objekt (Select-Object …) så kan resultatet exporteras till Excel efteråt.
Write-Host "Hej $Namn"
"""


def new_recipe():
    print()
    say("Ett recept är en vanlig .ps1 med ett litet huvud överst. Jag skapar en mall åt dig.", "happy")
    try:
        name = ask_text("Vad ska receptet heta?", required=True)
        cat = choose([dict(label=x, value=x) for x in ["SharePoint", "Exchange", "Teams", "Graph", "Excel", "Lokalt"]],
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
        ed = shutil.which("code") or None
        subprocess.run([ed, str(f)] if ed else ["open", "-t", str(f)])


# ───────────────────────────── profiler & guide ─────────────────────────────
GUID = re.compile(r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$")
PERMS = {
    "read": ("Bara läsa", "Sites.Read.All · Group.Read.All · User.Read.All",
             "-SharePointApplicationPermissions Sites.Read.All -GraphApplicationPermissions Group.Read.All,User.Read.All"),
    "full": ("Läsa och ändra", "Sites.FullControl.All · Group.ReadWrite.All · User.Read.All",
             "-SharePointApplicationPermissions Sites.FullControl.All -GraphApplicationPermissions Group.ReadWrite.All,User.Read.All"),
    "pnp": ("PnP:s standard", "det PnP.PowerShell väljer själv", ""),
}


def tenant_info(t):
    t = t.strip().lower()
    t = re.sub(r"^https?://", "", t).split("/")[0]
    if t.endswith(".sharepoint.com"):
        pre = t.split(".")[0].replace("-admin", "")
        return pre + ".onmicrosoft.com", pre
    if "." not in t:
        return t + ".onmicrosoft.com", t
    if t.endswith(".onmicrosoft.com"):
        return t, t.split(".")[0]
    return t, ""


def check_tenant(v):
    if " " in v.strip():
        return "Inga mellanslag. T.ex. contoso, contoso.onmicrosoft.com eller contoso.se."
    return None


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


def guide():
    """Ny PnP-profil: kund → tenant → URL:er → inloggningssätt → app (registrera / befintlig / senare) → spara."""
    print()
    say("Vi sätter upp PnP för en kund. Jag frågar, du svarar, och inget skapas i tenanten utan att du säger ja.",
        "happy", sub="Ctrl+C avbryter när som helst")
    try:
        names = {p["name"].lower() for p in profiles()}
        name = ask_text("Vad heter kunden? (profilnamn)", required=True,
                        check=lambda v: "Det namnet finns redan." if v.lower() in names else None)
        traw = ask_text("Kundens tenant? (contoso, contoso.onmicrosoft.com eller domän)", required=True,
                        check=check_tenant)
        tenant, prefix = tenant_info(traw)
        if not prefix:
            prefix = ask_text("SharePoint-namnet? (det som står före .sharepoint.com)", required=True,
                              check=check_tenant)
            prefix = prefix.lower().replace(".sharepoint.com", "").replace("-admin", "")
        admin_url = ask_text("Admin-URL", f"https://{prefix}-admin.sharepoint.com")
        root_url = f"https://{prefix}.sharepoint.com"
        auth = choose([dict(label=AUTHS["interactive"], hint="du loggar in, delegerade rättigheter", value="interactive"),
                       dict(label=AUTHS["device"], hint="kod i webbläsaren, bra över SSH", value="device"),
                       dict(label=AUTHS["certificate"], hint="obevakat, certifikat i Keychain", value="certificate")],
                      title="Hur ska du logga in?")
        if auth is None:
            raise Cancel()
        how = choose([dict(label="Registrera en ny app åt mig", hint="Register-PnPEntraID… (frågar först)", value="new"),
                      dict(label="Kunden har redan en app", hint="jag klistrar in ClientId", value="have"),
                      dict(label="Fyll i senare", hint="spara profilen utan ClientId", value="later")],
                     title="PnP kräver en egen Entra ID-app per tenant")
        if how is None:
            raise Cancel()
        prof = dict(name=name, tenant=tenant, prefix=prefix, admin_url=admin_url.rstrip("/"), root_url=root_url,
                    auth=auth, client_id="", upn="", created=datetime.now().isoformat(timespec="seconds"))
        if how == "have":
            prof["client_id"] = ask_text("ClientId (Application ID, en GUID)", required=True,
                                         check=lambda v: None if GUID.match(v) else "Det ser inte ut som en GUID.")
            if auth == "certificate":
                if not import_cert(prof):
                    raise Cancel()
        elif how == "new":
            if not register_app(prof):
                if not confirm("Spara profilen ändå (utan ClientId)?", True):
                    raise Cancel()
        prof["upn"] = ask_text("Ditt admin-konto hos kunden? (för Exchange, valfritt)", "")
    except Cancel:
        out(dim("  Avbrutet. Inget sparat."))
        return None
    ps = profiles()
    ps.append(prof)
    save_profiles(ps)
    set_config(active=prof["name"])
    WORKER.restart() if WORKER.conn else None
    print()
    celebrate(f"Profilen {prof['name']} är sparad och aktiv!")
    print_profile(prof)
    if prof["client_id"] and confirm("Testa anslutningen nu?", True):
        test_profile(prof)
    return prof


def import_cert(prof):
    path = ask_text("Sökväg till .pfx-filen", required=True,
                    check=lambda v: None if Path(os.path.expanduser(v)).is_file() else "Hittar ingen fil där.")
    pw = ask_text("Lösenordet till .pfx-filen (syns inte, sparas i Keychain)", secret=True)
    CERTS.mkdir(parents=True, exist_ok=True)
    os.chmod(CERTS, 0o700)
    dst = CERTS / (slug(prof["name"]) + ".pfx")
    shutil.copyfile(os.path.expanduser(path), dst)
    os.chmod(dst, 0o600)
    if not kc_set(prof["name"], pw):
        say("Kunde inte spara lösenordet i Keychain.", "sad")
        return False
    prof["cert_path"] = str(dst)
    out("  " + green("✔") + f" Certifikatet kopierat till {dst} · lösenordet i Keychain")
    return True


def register_app(prof):
    app = ask_text("Appens namn i Entra ID", f"PnP PowerShell ({prof['name']})")
    perm = "read"
    if prof["auth"] == "certificate":
        perm = choose([dict(label=v[0], hint=v[1], value=k) for k, v in PERMS.items()], title="Rättigheter för appen")
        if perm is None:
            raise Cancel()
    outdir = CERTS / ("ny-" + slug(prof["name"]))
    code = register_code(app, prof["tenant"], prof["auth"], perm, outdir)
    print()
    show_command(code, "Registrerar app")
    say(f"Det här skapar en app-registrering i {prof['tenant']} och kräver en admin som godkänner (samtycke). "
        "Ett webbläsarfönster öppnas.", "worry")
    if DRY or not confirm(f"Registrera appen i {prof['tenant']} på riktigt?", False):
        say("Torrkörning: ingen app skapad. Du kan köra kommandot själv och fylla i ClientId under Profiler.", "happy")
        return False
    secret = secrets.token_urlsafe(24) if prof["auth"] == "certificate" else None
    if secret:
        CERTS.mkdir(parents=True, exist_ok=True)
        os.chmod(CERTS, 0o700)
        outdir.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix="puppa-", suffix=".json")
    os.close(fd)
    env = dict(os.environ, PUPPA_OUT=tmp)
    if secret:
        env["PUPPA_SECRET"] = secret
    out(dim("  Lämnar över till PowerShell. Följ instruktionerna där…"))
    rc = subprocess.call(["pwsh", "-NoLogo", "-NoProfile", "-Command", REG_WRAP.format(cmd=code)], env=env)
    data = load_json(tmp, {})
    os.unlink(tmp)
    cid = next((v for k, v in data.items() if re.search(r"clientid|appid", k, re.I) and GUID.match(str(v))), "")
    if rc != 0 or not cid:
        say("Registreringen gick inte igenom (eller så hittade jag inget ClientId).", "sad")
        return False
    prof["client_id"] = cid
    if secret:
        pfx = next(iter(outdir.glob("*.pfx")), None)
        if not pfx:
            say("Appen skapades men jag hittar ingen .pfx-fil.", "sad")
            return False
        dst = CERTS / (slug(prof["name"]) + ".pfx")
        shutil.move(str(pfx), dst)
        os.chmod(dst, 0o600)
        for f in outdir.glob("*.cer"):
            shutil.move(str(f), CERTS / (slug(prof["name"]) + ".cer"))
        shutil.rmtree(outdir, ignore_errors=True)
        if not kc_set(prof["name"], secret):
            say("Kunde inte spara certifikatlösenordet i Keychain.", "sad")
            return False
        prof["cert_path"] = str(dst)
    out("  " + green("✔") + f" App registrerad · ClientId {cid}")
    return True


def print_profile(p):
    w = width()
    rows = [("Tenant", p["tenant"]), ("Admin-URL", p["admin_url"]), ("Rot-URL", p["root_url"]),
            ("Inloggning", AUTHS.get(p.get("auth"), "?")), ("ClientId", p.get("client_id") or yellow("saknas")),
            ("Exchange-konto", p.get("upn") or dim("–"))]
    if p.get("auth") == "certificate":
        rows.append(("Certifikat", (p.get("cert_path") or "–") + dim("  · lösenord i Keychain")))
    kw = 15 if w >= 58 else 0
    for k, v in rows:
        if kw:
            out("    " + dim(pad(k, kw)) + v)
        else:
            out("    " + dim(k))
            out("      " + v)


def profile_rows():
    act = config().get("active")
    rows = []
    w = width()
    for p in profiles():
        hint = p["tenant"] + ("  · " + AUTHS.get(p.get("auth"), "").split(" (")[0].lower() if w >= 90 else "")
        badge = green("aktiv") if p["name"] == act else (yellow("ofullst.") if not p.get("client_id") else "")
        rows.append(dict(label=p["name"], hint=hint, value=p["name"], badge=badge))
    return rows


def use_profile(p, quiet=False):
    if config().get("active") != p["name"]:
        set_config(active=p["name"])
        if WORKER.conn:
            WORKER.stop()
            WORKER.warm()
    if not quiet:
        say(f"Nu jobbar vi i {p['name']} ({p['tenant']}).", "proud")


def test_profile(p):
    print()
    if not p.get("client_id"):
        say(f"{p['name']} saknar ClientId, så det finns inget att testa än.", "sad")
        return False
    if config().get("active") != p["name"]:
        use_profile(p, quiet=True)
    code = connect_code("pnp", p, p["admin_url"]) + "\n" + \
        "$c = Get-PnPConnection\n[pscustomobject]@{ Url = $c.Url; Typ = \"$($c.ConnectionType)\"; " \
        "ClientId = $c.ClientId; Tenant = (Get-PnPTenantId) }"
    show_command(code, "Testar")
    if DRY:
        return True
    try:
        secret = need_secret("pnp", p)
    except RuntimeError as e:
        say(str(e), "sad")
        return False
    ok, n, _ = run_ps(f"Testar {p['name']}", code, secret)
    if ok:
        WORKER.conn["pnp"] = (p["name"], p["admin_url"])
        celebrate(f"{p['name']} fungerar! Du är ansluten till SharePoint-admin.")
    else:
        say("Anslutningen misslyckades. Kolla ClientId, tenant och att appen har samtycke.", "sad")
    return ok


def delete_profile(p):
    say(f"Ta bort {p['name']}? Profilen" + (", certifikatet och lösenordet i Keychain" if p.get("cert_path") else "")
        + " försvinner härifrån. Appen i Entra ID finns kvar.", "worry")
    if not confirm(f"Ta bort {p['name']}?", False):
        return False
    save_profiles([x for x in profiles() if x["name"] != p["name"]])
    kc_del(p["name"])
    for ext in (".pfx", ".cer"):
        f = CERTS / (slug(p["name"]) + ext)
        if f.exists():
            f.unlink()
    if config().get("active") == p["name"]:
        rest = profiles()
        set_config(active=rest[0]["name"] if rest else None)
        WORKER.stop()
        WORKER.warm()
    say(f"{p['name']} är borttagen.", "happy")
    return True


def set_client_id(p):
    try:
        cid = ask_text("ClientId (Application ID)", p.get("client_id", ""), required=True,
                       check=lambda v: None if GUID.match(v) else "Det ser inte ut som en GUID.")
        if p.get("auth") == "certificate" and not (p.get("cert_path") and kc_get(p["name"])):
            if not import_cert(p):
                return
    except Cancel:
        out(dim("  Avbrutet."))
        return
    ps = profiles()
    for x in ps:
        if x["name"] == p["name"]:
            x.update(client_id=cid, cert_path=p.get("cert_path", x.get("cert_path", "")))
    save_profiles(ps)
    say("ClientId sparat.", "proud")


def open_shell(p):
    """Lämnar över terminalen till ett pwsh-skal som redan är anslutet till kundens SharePoint-admin."""
    if not p.get("client_id"):
        say(f"{p['name']} saknar ClientId.", "sad")
        return
    code = connect_code("pnp", p, p["admin_url"])
    print()
    show_command(code, "Öppnar skal")
    if DRY:
        return
    try:
        secret = need_secret("pnp", p)
    except RuntimeError as e:
        say(str(e), "sad")
        return
    env = dict(os.environ)
    if secret:
        env["PUPPA_SECRET"] = secret
    script = code + "\nRemove-Item Env:PUPPA_SECRET -ErrorAction SilentlyContinue\n" + \
        f"Write-Host \"Ansluten till {p['name']} – skriv exit för att gå tillbaka till Puppa.\" -ForegroundColor Cyan"
    say(f"Öppnar PowerShell anslutet till {p['name']}. Skriv exit för att komma tillbaka.", "happy")
    subprocess.call(["pwsh", "-NoLogo", "-NoExit", "-Command", script], env=env)


def profiles_menu():
    while True:
        ps = profiles()
        if not ps:
            say("Inga profiler än. Ska vi skapa en?", "worry")
            if confirm("Starta guiden?", True):
                guide()
            return
        rows = profile_rows() + [("head", ""), dict(label="Ny profil…", value="__new", key="n"),
                                   dict(label="Tillbaka", value=None, key="q")]
        name = choose(rows, title="Profiler", footer="enter väljer · n ny · q tillbaka", screen=True)
        if name == "__new":
            guide()
            continue
        if not name:
            return
        p = find_profile(name)
        out(f"  {accent('▍')}{bold(p['name'])}")
        print_profile(p)
        act = choose([dict(label="Använd", hint="gör aktiv", value="use"),
                      dict(label="Testa anslutningen", hint="Connect-PnPOnline mot admin", value="test"),
                      dict(label="Öppna PowerShell-skal", hint="anslutet, du tar över", value="shell"),
                      dict(label="Ange ClientId…", hint="om appen fanns/skapades utanför Puppa", value="cid"),
                      dict(label="Ta bort…", value="del"),
                      dict(label="Tillbaka", value=None, key="q")], title=p["name"])
        if act == "use":
            use_profile(p, quiet=True)
            return
        if act == "test":
            test_profile(p)
        elif act == "shell":
            open_shell(p)
        elif act == "cid":
            set_client_id(p)
        elif act == "del":
            delete_profile(p)


LOGOUT_PS = r"""
$ut = [System.Collections.Generic.List[string]]::new()
try { Import-Module PnP.PowerShell -ErrorAction Stop; Disconnect-PnPOnline -ClearPersistedLogin -ErrorAction Stop; $ut.Add('PnP (även sparad inloggning)') } catch { }
try { Import-Module ExchangeOnlineManagement -ErrorAction Stop; Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop *> $null; $ut.Add('Exchange Online') } catch { }
try { Import-Module MicrosoftTeams -ErrorAction Stop; Disconnect-MicrosoftTeams -Confirm:$false -ErrorAction Stop *> $null; $ut.Add('Teams') } catch { }
try { Import-Module Microsoft.Graph.Authentication -ErrorAction Stop; Disconnect-MgGraph -ErrorAction Stop *> $null; $ut.Add('Microsoft Graph (även sparad token)') } catch { }
foreach ($x in $ut) { Write-Host "  utloggad: $x" }
"""
LOGOUT_SHOW = """Disconnect-PnPOnline -ClearPersistedLogin
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-MicrosoftTeams -Confirm:$false
Disconnect-MgGraph"""


def logout():
    """Loggar ut från PnP, Exchange, Teams och Graph, rensar sparade inloggningar och startar en ren session."""
    print()
    show_command(LOGOUT_SHOW, "Loggar ut")
    if DRY:
        return True
    ok, _, lines = run_ps("Loggar ut från allt", LOGOUT_PS, show_output=False)
    done = [t.split("utloggad: ", 1)[1] for k, t in lines if k == "H" and "utloggad: " in t]
    WORKER.stop()
    WORKER.warm()
    for d in done:
        out("  " + green("✔") + " " + d)
    say("Utloggad från allt. Nästa recept frågar efter inloggning igen. "
        "(Webbläsarens Microsoft-inloggning ligger kvar – logga ut där också om datorn delas.)", "happy")
    return True


def connect_menu():
    p = active_profile()
    if not p:
        say("Välj eller skapa en profil först.", "worry")
        return
    rows = [dict(label="SharePoint-admin (PnP)", hint=p["admin_url"], value="pnp-admin"),
            dict(label="SharePoint-sajt (PnP)…", hint="valfri sajt-URL", value="pnp-site"),
            dict(label="Exchange Online", hint="Connect-ExchangeOnline", value="exo"),
            dict(label="Teams", hint="Connect-MicrosoftTeams", value="teams"),
            dict(label="Microsoft Graph…", hint="Connect-MgGraph med valda scopes", value="graph"),
            dict(label="Logga ut från allt", hint="PnP · EXO · Teams · Graph + sparade inloggningar", value="off"),
            dict(label="Tillbaka", value=None, key="q")]
    v = choose(rows, title=f"Anslut · {p['name']}", status=conn_status, screen=True)
    try:
        if v == "pnp-admin":
            ensure("pnp", p, p["admin_url"])
        elif v == "pnp-site":
            url = ask_text("Sajtens URL", p["root_url"] + "/sites/", required=True)
            ensure("pnp", p, url)
        elif v in ("exo", "teams"):
            ensure(v, p)
        elif v == "graph":
            sc = ask_text("Scopes (kommaseparerade)", "User.Read.All,Group.Read.All,Organization.Read.All")
            ensure("graph", p, scopes=[s.strip() for s in sc.split(",") if s.strip()])
        elif v == "off":
            logout()
    except Cancel:
        out(dim("  Avbrutet."))


CAT_HINTS = {"SharePoint": "sajter, behörigheter, papperskorg, delning", "OneDrive": "hitta, ge åtkomst, största",
             "Exchange": "delade postlådor, åtkomst, kalender, listor, spårning",
             "Grupper": "medlemmar, lägg till/ta bort, utan ägare", "Användare": "info, licenser, gäster, MFA, spärra",
             "Teams": "team, medlemmar, kanaler", "Excel": "exportera senaste resultatet", "Lokalt": "moduler, session"}


def remember_recipe(rid):
    rec = [x for x in config().get("recent", []) if x != rid]
    set_config(recent=([rid] + rec)[:5])


def recipe_rows(rs):
    rows = [dict(label=r.name, hint=r.desc, value=r.id, badge=r.badge()) for r in rs if not r.writes]
    w = [dict(label=r.name, hint=r.desc, value=r.id, badge=r.badge()) for r in rs if r.writes]
    if rows and w:
        rows.append(("head", "ändrar"))
    return rows + w


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
        cat = choose(rows, title="Recept · välj område", start=cat_last, status=conn_status, screen=True)
        if cat is None:
            return
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
    ps = profiles()
    p = active_profile()
    if not ps:
        return "Hej! Jag är Puppa. Du har inga PnP-profiler än – ska vi skapa den första?", "worry"
    if not p:
        return "Välj vilken kund du jobbar med i dag.", "happy"
    if not p.get("client_id"):
        return f"{p['name']} saknar fortfarande ClientId. Fyll i det under Profiler, eller kör ett lokalt recept.", "worry"
    if not WORKER.conn:
        return f"Hej! Vi jobbar i {p['name']}. Kör ett recept – jag ansluter när det behövs.", "happy"
    return f"Ansluten och redo i {p['name']}. Vad blir nästa steg?", "happy"


def main_menu():
    last = None
    while True:
        msg, mood = greeting()
        ps = profiles()
        p = active_profile()
        rows = [dict(label="Recept…", hint="kör en uppgift: lista, skapa, exportera", value="recipes"),
                dict(label="Byt kund", hint=(p["name"] if p else "välj profil"), value="switch"),
                dict(label="Ny PnP-profil", hint="guide: app + profil på några steg", value="new"),
                dict(label="Profiler…", hint="lista, testa, ClientId, ta bort", value="profiles"),
                dict(label="Anslut…", hint="PnP · Exchange · Teams · Graph", value="connect"),
                dict(label="PowerShell-skal", hint="anslutet till kunden, du tar över", value="shell"),
                dict(label="Logga ut från allt", hint="PnP · Exchange · Teams · Graph", value="logout", key="u"),
                dict(label="Avsluta", value="quit", key="q")]
        if not ps:
            rows.insert(0, rows.pop(2))
            rows = [r for r in rows if r["value"] not in ("profiles", "connect", "shell")]
        if len(ps) < 2:
            rows = [r for r in rows if r["value"] != "switch"]
        v = choose(rows, msg=msg, mood=mood, status=conn_status, start=last,
                   footer="↑/↓ · enter · siffra väljer · q avslutar" if width() >= 50 else None,
                   sub=f"puppa {VERSION}" + ("  ·  torrkörning" if DRY else ""), screen=True)
        last = v
        if v in (None, "quit"):
            print()
            if WORKER.conn and confirm("Logga ut från allt innan du går?", False):
                logout()
            say(random.choice(["Hej då!", "Vi ses!",
                               "Hej då! -WhatIf var med dig."]), "happy")
            print()
            return
        if v == "recipes":
            recipes_menu()
        elif v == "switch":
            name = choose(profile_rows() + [dict(label="Tillbaka", value=None, key="q")], title="Byt kund",
                          start=config().get("active"), screen=True)
            if name:
                use_profile(find_profile(name), quiet=True)
        elif v == "new":
            guide()
        elif v == "logout":
            logout()
        elif v == "profiles":
            profiles_menu()
        elif v == "connect":
            connect_menu()
        elif v == "shell":
            if p:
                open_shell(p)
            else:
                say("Välj eller skapa en profil först.", "worry")


# ───────────────────────────── CLI ─────────────────────────────
def list_profiles():
    ps = profiles()
    if not ps:
        say("Inga profiler än. Kör puppa ny.", "worry")
        return
    act = config().get("active")
    w = width()
    nw = min(max(len(p["name"]) for p in ps) + 2, 24)
    for p in ps:
        mark = green("●") if p["name"] == act else dim("○")
        cid = (p.get("client_id") or "")[:8] or yellow("saknas")
        if w >= 72:
            out(f"  {mark} {pad(bold(cut(p['name'], nw)), nw)}{pad(cut(p['tenant'], 30), 32)}"
                f"{pad(AUTHS.get(p.get('auth'), '?').split(' (')[0], 13)}{dim(cid)}")
        elif w >= 44:
            out(f"  {mark} {pad(bold(cut(p['name'], nw)), nw)}{dim(cut(p['tenant'], w - nw - 6))}")
        else:
            out(f"  {mark} {bold(p['name'])}")
            out(f"    {dim(cut(p['tenant'], w - 6))}")


def list_recipes():
    rs = load_recipes()
    w = width()
    nw = min(max(len(r.id) for r in rs) + 2, 30)
    cat = None
    for r in rs:
        if r.category != cat:
            cat = r.category
            out("  " + accent("▍") + bold(cat))
        if w >= 58:
            out(f"    {pad(cut(r.id, nw), nw)}{pad(r.badge(), 8)}" + (dim(r.name) if w >= 72 else ""))
        else:
            out(f"    {cut(r.id, w - 14)} {r.badge()}")


def main(argv=None):
    import argparse
    global ANIM, DRY, ITTY
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
    sub.add_parser("ny", help="guide: ny PnP-profil")
    sub.add_parser("profiler", help="lista profiler")
    u = sub.add_parser("använd", aliases=["anvand", "use"], help="gör en profil aktiv")
    u.add_argument("profil")
    t = sub.add_parser("test", help="testa PnP-anslutning")
    t.add_argument("profil", nargs="?")
    s = sub.add_parser("skal", aliases=["shell"], help="öppna anslutet pwsh-skal")
    s.add_argument("profil", nargs="?")
    sub.add_parser("recept", help="lista recept")
    k = sub.add_parser("kör", aliases=["kor", "run"], help="kör ett recept")
    k.add_argument("recept")
    k.add_argument("par", nargs="*", help="Namn=värde")
    k.add_argument("--ja", action="store_true", help="bekräfta ändrande recept utan att fråga")
    sub.add_parser("logga-ut", aliases=["logout"], help="logga ut från PnP, Exchange, Teams och Graph")
    sub.add_parser("hej")
    a = ap.parse_args(rest)

    def sigint(*_):
        sys.stdout.write("\033[?25h\n")
        sys.stdout.flush()
        WORKER.stop()
        os._exit(130)

    if not shutil.which("pwsh"):
        say("Jag hittar inte pwsh (PowerShell 7). Installera med: brew install powershell", "sad")
        return 1

    def pick_profile(name):
        p = find_profile(name) if name else active_profile()
        if not p:
            say(f"Hittar ingen profil {name or '(ingen aktiv)'}. Se puppa profiler.", "sad")
            sys.exit(1)
        return p

    try:
        if a.cmd is None:
            if not ITTY or not TTY:
                print(__doc__)
                return 0
            signal.signal(signal.SIGINT, signal.default_int_handler)
            WORKER.warm()
            print()
            main_menu()
            return 0
        if a.cmd == "ny":
            if not ITTY:
                sys.exit("  guiden behöver en riktig terminal")
            WORKER.warm()
            return 0 if guide() else 1
        if a.cmd == "profiler":
            list_profiles()
            return 0
        if a.cmd in ("använd", "anvand", "use"):
            use_profile(pick_profile(a.profil))
            return 0
        if a.cmd == "test":
            WORKER.warm()
            return 0 if test_profile(pick_profile(a.profil)) else 1
        if a.cmd in ("skal", "shell"):
            open_shell(pick_profile(a.profil))
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
    except KeyboardInterrupt:
        sigint()
    finally:
        WORKER.stop()
    return 0


if __name__ == "__main__":
    sys.exit(main())

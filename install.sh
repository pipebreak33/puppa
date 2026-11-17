#!/usr/bin/env sh
# Puppa för macOS och Linux – installera eller uppdatera (kör igen = senaste versionen):
#   curl -fsSL https://raw.githubusercontent.com/pipebreak33/puppa/main/install.sh | sh
# Från en lokal kopia:  ./install.sh
# Styr källan med PUPPA_REPO=ägare/repo och PUPPA_BRANCH=gren (eller en release-tagg).
# Inga frågor: exitkod 0 = installerat, annat = fel.
set -eu

REPO="${PUPPA_REPO:-pipebreak33/puppa}"
BRANCH="${PUPPA_BRANCH:-main}"
DIR="$HOME/.puppa"
APP="$DIR/app"
BIN="$HOME/.local/bin"

say()  { printf '  \033[38;5;75mP›\033[0m %s\n' "$1"; }
warn() { printf '  \033[38;5;214m!\033[0m %s\n' "$1"; }
die()  { printf '  \033[31m✘\033[0m %s\n' "$1" >&2; exit 1; }

# 1. Python 3.8+ (Puppa är ett Python-skript utan beroenden)
PY=""
for c in python3 python; do
  if command -v "$c" >/dev/null 2>&1 && "$c" -c 'import sys; sys.exit(sys.version_info < (3, 8))' 2>/dev/null; then
    PY="$(command -v "$c")"; break
  fi
done
if [ -z "$PY" ]; then
  case "$(uname -s)" in
    Darwin) die "Python 3.8+ saknas. Kör: xcode-select --install   (eller: brew install python)";;
    *)      die "Python 3.8+ saknas. Kör: sudo apt install python3   (eller dnf/pacman motsvarande)";;
  esac
fi

# 2. Källan: en lokal checkout bredvid skriptet, annars arkivet från GitHub
SRC=""
case "$0" in
  *install.sh)
    HERE="$(cd "$(dirname "$0")" 2>/dev/null && pwd || true)"
    if [ -n "$HERE" ] && [ -f "$HERE/puppa.py" ] && [ -d "$HERE/recipes" ] && [ "${PUPPA_FROM_WEB:-}" != "1" ]; then
      SRC="$HERE"
      say "Installerar från $HERE"
    fi;;
esac
TMP=""
cleanup() { if [ -n "$TMP" ]; then rm -rf "$TMP"; fi; return 0; }
trap cleanup EXIT
if [ -z "$SRC" ]; then
  command -v curl >/dev/null 2>&1 || die "curl saknas."
  command -v tar >/dev/null 2>&1 || die "tar saknas."
  TMP="$(mktemp -d)"
  say "Hämtar Puppa från github.com/$REPO ($BRANCH)"
  curl -fsSL "https://github.com/$REPO/archive/$BRANCH.tar.gz" -o "$TMP/puppa.tar.gz" \
    || die "nedladdningen misslyckades – kolla nätet eller PUPPA_REPO/PUPPA_BRANCH"
  tar -xzf "$TMP/puppa.tar.gz" -C "$TMP" || die "kunde inte packa upp arkivet"
  SRC="$(find "$TMP" -maxdepth 2 -name puppa.py -exec dirname {} \; | head -n 1)"
  [ -n "$SRC" ] && [ -d "$SRC/recipes" ] || die "arkivet innehåller inte Puppa"
fi
"$PY" -c 'import ast,sys; ast.parse(open(sys.argv[1],encoding="utf-8").read())' "$SRC/puppa.py" \
  || die "puppa.py är trasig"

# 3. Byt ut programmet (~/.puppa/app). Dina tenants, inställningar och egna recept i ~/.puppa rörs inte.
mkdir -p "$DIR" "$BIN"
chmod 700 "$DIR" 2>/dev/null || true
rm -rf "$APP.new"
mkdir -p "$APP.new"
cp "$SRC/puppa.py" "$APP.new/puppa.py"
cp -R "$SRC/recipes" "$APP.new/recipes"
for f in README.md LICENSE; do [ -f "$SRC/$f" ] && cp "$SRC/$f" "$APP.new/$f"; done
rm -rf "$APP"
mv "$APP.new" "$APP"

# 4. Startkommando i ~/.local/bin (en ev. gammal symlänk tas bort först – skriv aldrig genom den)
rm -f "$BIN/puppa"
cat > "$BIN/puppa" <<SHIM
#!/usr/bin/env sh
exec "$PY" "$APP/puppa.py" "\$@"
SHIM
chmod +x "$BIN/puppa"

# 5. ~/.local/bin i PATH
case ":$PATH:" in
  *":$BIN:"*) ;;
  *)
    case "${SHELL:-}" in
      */zsh)  RC="$HOME/.zshrc" ;;
      */bash) RC="$HOME/.bashrc"; [ "$(uname -s)" = Darwin ] && RC="$HOME/.bash_profile" ;;
      *)      RC="$HOME/.profile" ;;
    esac
    grep -qs '.local/bin' "$RC" 2>/dev/null || printf '\nexport PATH="$HOME/.local/bin:$PATH"\n' >> "$RC"
    say "La till ~/.local/bin i PATH ($RC). Öppna en ny terminal."
    ;;
esac

# 6. Förutsättningar: PowerShell 7 och modulerna (bara varningar – Puppa är installerat ändå)
if command -v pwsh >/dev/null 2>&1; then
  MISSING="$(pwsh -NoLogo -NoProfile -NonInteractive -Command '@("PnP.PowerShell","ExchangeOnlineManagement","MicrosoftTeams","Microsoft.Graph.Authentication","Microsoft.PowerShell.SecretManagement","Microsoft.PowerShell.SecretStore","ImportExcel") | Where-Object { -not (Get-Module -ListAvailable -Name $_) }' 2>/dev/null | tr '\n' ' ' || true)"
  if [ -n "$(printf '%s' "$MISSING" | tr -d ' ')" ]; then
    warn "Moduler som saknas: $MISSING"
    warn "Installera: pwsh -c 'Install-PSResource $(printf '%s' "$MISSING" | sed 's/ *$//; s/ /,/g') -Scope CurrentUser -TrustRepository'   (eller: uppa install m365)"
  fi
else
  case "$(uname -s)" in
    Darwin) warn "PowerShell 7 (pwsh) saknas. Installera: brew install powershell   (eller: uppa install powershell)";;
    *)      warn "PowerShell 7 (pwsh) saknas. Se https://aka.ms/install-powershell   (eller: uppa install powershell)";;
  esac
fi

VER="$("$PY" "$APP/puppa.py" --version 2>/dev/null || echo puppa)"
printf '\n  \033[38;5;42m✔\033[0m %s installerad. Kör: \033[1mpuppa\033[0m\n\n' "$VER"

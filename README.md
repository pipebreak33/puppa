# Puppa

Vanliga M365-adminuppgifter i PowerShell med några knapptryck, i stället för att skriva cmdlets.
Syskon till [uppa](https://github.com/pipebreak33/uppa). Maskoten är ett P med ögon: den föreslår nästa steg, blinkar och tittar runt
medan något körs, blir glad när det gick bra och ledsen vid fel.

## Installera

Kräver macOS (Keychain), Python 3.8+, PowerShell 7 (`brew install powershell`) och de moduler du vill använda:
PnP.PowerShell 3.x, ExchangeOnlineManagement, MicrosoftTeams, Microsoft.Graph, ImportExcel.

```bash
git clone https://github.com/pipebreak33/puppa.git ~/puppa
```

```bash
ln -sf ~/puppa/puppa.py ~/.local/bin/puppa
```

```bash
puppa
```

## Kom igång

1. **Ny PnP-profil**: guiden frågar efter kund, tenant och inloggningssätt och registrerar sedan
   Entra ID-appen (`Register-PnPEntraIDAppForInteractiveLogin` eller `Register-PnPEntraIDApp` för
   certifikat). Du ser exakt vilket kommando som körs, och inget skapas i tenanten om du inte svarar **j**.
   Har kunden redan en app klistrar du in dess ClientId i stället.
2. **Recept…** → välj område → välj recept. Puppa frågar efter det som behövs, visar kommandot och
   ansluter själv (PnP/Exchange/Teams/Graph) första gången. Anslutningarna lever kvar under sessionen.
3. Recept märkta **ändrar** kräver alltid ett **j**. Recept märkta **läs** körs direkt.
4. **Exportera senaste resultat till Excel/CSV** sparar förra receptets resultat i `~/Documents/Puppa/`.
5. **Logga ut från allt** (`u`) kopplar från PnP, Exchange, Teams och Graph och rensar sparade inloggningar.

## Kommandon

| | |
|---|---|
| `puppa` | menyn |
| `puppa ny` | guiden för ny PnP-profil |
| `puppa profiler` · `puppa använd KUND` · `puppa test [KUND]` | profiler |
| `puppa skal [KUND]` | pwsh-skal som redan kört `Connect-PnPOnline`, du tar över |
| `puppa recept` | lista alla recept |
| `puppa kör RECEPT Namn=värde … [--ja]` | kör ett recept utan meny |
| `puppa logga-ut` | logga ut från allt |
| `-n` / `--torr` | visa kommandona, kör ingenting |
| `--no-anim` | inga animationer (stängs också av av `NO_COLOR` och när utdata pipas) |

## Recept (63 st)

SharePoint (sajter, största sajter, sajtinfo, behörigheter, användare, listor, papperskorg/återställ,
extern delning, webbplatsadmins, SharePoint-grupper, raderade sajter, tenantinställningar) ·
OneDrive (hitta, största, ge åtkomst) · Exchange (postlådor, delade postlådor: skapa, ge/ta bort åtkomst,
vem har åtkomst, vad kommer en användare åt; konvertera till delad, vidarebefordringar, autosvar,
kalenderbehörighet, rum, distributionslistor, meddelandespårning) · Grupper (lista, medlemmar,
användarens grupper, lägg till/ta bort, utan ägare) · Användare (info, licenser, vem har licens,
inaktiva, gäster, bjud in gäst, spärra inloggning, MFA-metoder, adminroller) · Teams (team, medlemmar,
lägg till, kanaler, team med en ägare) · Excel/CSV-export · Lokalt (modulversioner, aktuell session).

### Eget recept

Välj **Nytt recept (mall)…** eller lägg en `.ps1` i `~/.puppa/recipes/` (den vinner över inbyggda med samma namn):

```powershell
<#PUPPA
namn:        Lista team
kategori:    Teams
typ:         läs                 # läs | ändrar | lokal
kräver:      teams               # ingen | pnp-admin | pnp:Url | exo | teams | graph:Scope1,Scope2
beskrivning: Alla team med synlighet.
param:       Filter? | Del av teamnamnet |                 # ? = valfri
param:       Roll | Roll | Member | Member,Owner           # namn | fråga | standard | val
#>
Get-Team | Where-Object { -not $Filter -or $_.DisplayName -like "*$Filter*" } | Select-Object DisplayName, Visibility
```

Parametrarna blir variabler (`$Filter`). I standardvärden och skript fungerar `{tenant}` `{prefix}`
`{adminUrl}` `{rootUrl}` `{upn}` `{profil}` `{datum}`. Returnera objekt så går resultatet att exportera.

## Säkerhet

- `~/.puppa/profiles.json` innehåller inga hemligheter (tenant, URL:er, ClientId, inloggningssätt).
- Certifikatprofiler: `.pfx` i `~/.puppa/certs/` (chmod 600, lösenordsskyddad) och lösenordet i
  **macOS Keychain** (tjänst `puppa`). Lösenordet skickas till PowerShell via en miljövariabel som tas
  bort direkt efter anslutningen, aldrig på kommandoraden.
- Interaktiv/enhetskod använder `-PersistLogin`, så du slipper logga in varje gång. **Logga ut från allt** rensar det.

## Teknik

En enda Python-fil (bara standardbiblioteket, 3.8+) plus recepten som `.ps1`. En långlivad `pwsh`-process
kör recepten så att anslutningarna överlever mellan körningar. Layouten skalar med terminalen: inga rader
radbryts, smal layout under 58 kolumner, maskoten döljs under 44, och allt ritas om vid fönsterändring.

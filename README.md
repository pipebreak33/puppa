# Puppa

<p align="center">
  <img src="assets/puppa.svg" alt="Puppa: Hej! Jag är Puppa." width="720">
</p>

Vanliga M365-adminuppgifter i PowerShell med några knapptryck, i stället för att skriva cmdlets.
Fungerar likadant på **Windows, macOS och Linux**. Syskon till [uppa](https://github.com/pipebreak33/uppa).
Maskoten är ett P med ögon: den föreslår nästa steg, blinkar och tittar runt medan något körs,
blir glad när det gick bra och ledsen vid fel.

## Installera (eller uppdatera – kör samma rad igen)

**Windows** (PowerShell):

```powershell
irm https://raw.githubusercontent.com/pipebreak33/puppa/main/install.ps1 | iex
```

**macOS / Linux:**

```bash
curl -fsSL https://raw.githubusercontent.com/pipebreak33/puppa/main/install.sh | sh
```

Installern hittar (eller på Windows installerar) Python 3.8+, lägger Puppa i `~/.puppa/app` och skapar kommandot
`puppa` (`~/.local/bin/puppa`, eller `%USERPROFILE%\.puppa\bin\puppa.cmd` + PATH på Windows). Dina tenants och
egna recept i `~/.puppa` rörs aldrig vid uppdatering. Den varnar om PowerShell 7 eller modulerna saknas.
Med [uppa](https://github.com/pipebreak33/uppa) kan Puppa också installeras och uppdateras därifrån.

Styr källan med `PUPPA_REPO=ägare/repo` och `PUPPA_BRANCH=gren`. Från en lokal kopia: `./install.sh` / `./install.ps1`.

Kräver PowerShell 7 (`pwsh`) och de moduler du vill använda: PnP.PowerShell 3.x, ExchangeOnlineManagement,
MicrosoftTeams, Microsoft.Graph, ImportExcel, Microsoft.PowerShell.SecretManagement + SecretStore.

## Så funkar det

- **Ingen förberedelse.** Välj ett recept. Puppa loggar in när det behövs (Exchange, Graph, Teams) och lägger
  tenanten i **Tenants**-listan med domän och konto, så att du kan växla mellan kunder senare.
- **Fortsätt där du lämnade.** Startskärmen visar tenanten du jobbade i och dina senaste recept överst.
- **Sök överallt.** Börja skriva i vilken meny som helst så filtreras den fuzzy. I huvudmenyn söker du direkt
  bland alla recept (`delad`, `postl`, `spärra` …).
- **Sök i tenanten.** När ett recept frågar efter användare, grupp, postlåda, distributionslista, sajt, team eller
  licens räcker en del av namnet – `johan` ger en lista att välja ur, en enda träff väljs direkt.
- **SharePoint = PnP.** Bara SharePoint-recepten går via PnP PowerShell, som kräver en egen Entra ID-app per tenant.
  Första gången frågar Puppa om en **ny PnP-appregistrering** (`Register-PnPEntraIDAppForInteractiveLogin` /
  `Register-PnPEntraIDApp`) eller ett befintligt ClientId. Inget skapas i tenanten utan att du svarar **j**.
- **Du ser alltid kommandot** som körs. Recept märkta **ändrar** kräver ett **j**, och Puppa loggar in först så att
  du ser vilken tenant ändringen hamnar i. **läs**-recept körs direkt.
- **Exportera** förra resultatet till Excel (ImportExcel) eller CSV (semikolon, UTF-8) i `~/Documents/Puppa/`.
- **Logga ut…** från allt (förvalt) eller en tjänst i taget. Rensar även sparade inloggningar
  (PnP `-ClearPersistedLogin`, Graph-token). Puppa frågar också när du avslutar.

## Kommandon

| | |
|---|---|
| `puppa` | menyn |
| `puppa tenants` · `puppa tenant NAMN` | lista / växla tenant |
| `puppa ny` | lägg till en tenant för hand (oftast onödigt) |
| `puppa pnp [TENANT]` | ny PnP-appregistrering (eller ange ClientId) |
| `puppa test [TENANT]` · `puppa skal [TENANT]` | testa PnP / pwsh-skal som redan kört `Connect-PnPOnline` |
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
namn:        Medlemmar i ett team
kategori:    Teams
typ:         läs                 # läs | ändrar | lokal
kräver:      teams               # ingen | pnp-admin | pnp:Url | exo | teams | graph:Scope1,Scope2
beskrivning: Ägare, medlemmar och gäster i ett team.
param:       Team | Teamets namn | | @team               # @… = sök i tenanten
param:       Roll | Roll | Member | Member,Owner          # namn | fråga | standard | val
param:       Filter? | Del av namnet |                    # ? = valfri
#>
$t = Get-Team -DisplayName $Team
Get-TeamUser -GroupId $t.GroupId | Where-Object { $_.Role -eq $Roll } | Select-Object Name, User, Role
```

Uppslag: `@user` `@group` `@sku` (Graph) · `@mailbox` `@recipient` `@dl` (Exchange) · `@site` `@spuser` (PnP) ·
`@team` (Teams). Lägg till `+` för flera värden (`@user+` blir `a@x.se,b@x.se`).
Parametrarna blir variabler (`$Team`). I standardvärden och skript fungerar `{tenant}` `{prefix}`
`{adminUrl}` `{rootUrl}` `{upn}` `{kund}` `{datum}`. Returnera objekt så går resultatet att exportera.

## Säkerhet

- `~/.puppa/kunder.json` innehåller inga hemligheter (namn, domän, tenant-id, konto, ClientId, inloggningssätt).
- Certifikatlösenord (app-only PnP) ligger i PowerShells **SecretManagement**, i valvet `Puppa` (SecretStore),
  på alla plattformar. Saknas ett valv frågar Puppa om det ska låsas med eget lösenord eller bara med ditt
  användarkonto. Lösenordet når PowerShell via en miljövariabel som tas bort direkt efter anslutningen,
  aldrig på kommandoraden. `.pfx`-filen ligger lösenordsskyddad i `~/.puppa/certs/`.

## Teknik

En Python-fil (bara standardbiblioteket, 3.8+) plus recepten som `.ps1`. En långlivad `pwsh`-process kör recepten,
så inloggningarna lever mellan körningarna. Tangenter läses med `msvcrt` på Windows och `termios` på macOS/Linux.
Layouten skalar med terminalen: inga rader radbryts, smal layout under 58 kolumner, maskoten döljs under 44, och allt
ritas om när fönstret ändrar storlek (även på Windows, där SIGWINCH saknas).

# M365-SPO-OD-Teams Interactive Backup

Zálohuje soubory z knihoven dokumentů SharePoint Online, dokumentů z kanálů Teams a OneDrive pro firmy do lokálního adresáře. Podporuje inkrementální zálohování, automatické řízení omezování (throttling) a režim zkušebního běhu (dry-run).

## Předpoklady

### PowerShell moduly

Nainstalujte požadované moduly Microsoft Graph PowerShell SDK:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module Microsoft.Graph.Files        -Scope CurrentUser
Install-Module Microsoft.Graph.Sites        -Scope CurrentUser
```

### Oprávnění

Skript se ověřuje interaktivně přes `Connect-MgGraph` a vyžaduje následující oprávnění (scopes):

| Oprávnění | Účel |
|---|---|
| `Files.Read.All` | Čtení souborů z libovolného disku (knihovny SPO, OneDrive) |
| `Sites.Read.All` | Vyhledávání SharePoint webů a jejich knihoven dokumentů |
| `User.Read` | Čtení profilu přihlášeného uživatele (nutné pro přístup k OneDrive přes `me/drive`) |

Správce Azure AD může potřebovat udělit souhlas (consent) pro tato oprávnění ve vašem tenantovi.

### Prostředí

- PowerShell 5.1 nebo PowerShell 7+
- Síťový přístup k Microsoft Graph (`graph.microsoft.com`)

## Rychlý start

1. Umístěte soubor `config.json` do stejného adresáře jako skript (viz Konfigurace níže).
2. Spusťte skript:

```powershell
# Náhled - co by se stáhlo (žádné soubory se nezapisují)
.\Backup-M365-Interactive.ps1 -DryRun

# Spuštění skutečné zálohy (výchozí: režim RenameNew — originály se zachovají)
.\Backup-M365-Interactive.ps1

# Spuštění zálohy v režimu přepsání (nahradí změněné soubory na místě)
.\Backup-M365-Interactive.ps1 -UpdateAction Overwrite
```

3. Při prvním spuštění se otevře okno prohlížeče pro ověření identity.

### Parametry

| Parametr | Výchozí | Popis |
|---|---|---|
| `-DryRun` | vypnuto | Náhled toho, co by se stáhlo, bez zápisu souborů. |
| `-UpdateAction` | `RenameNew` | Globální režim aktualizace. `RenameNew` přejmenuje existující lokální soubor s příponou `_prev_XXXXX` a stáhne novou verzi pod původním názvem (zachová starý soubor a zároveň zajistí správné inkrementální porovnání). `Overwrite` přepíše existující soubory na místě. Lze přepsat per-úlohu pomocí pole `UpdateAction` v `config.json`. |

## Konfigurace

Vytvořte soubor `config.json` v adresáři skriptu. Jedná se o JSON pole objektů úloh.

### Úloha SharePoint / Teams

```json
[
  {
    "Type": "SharePoint",
    "SiteName": "Marketing",
    "LibraryName": "Shared Documents",
    "SourcePath": "/",
    "LocalDownloadPath": "C:\\Backups\\Marketing",
    "UpdateAction": "Overwrite"
  }
]

> **Poznámka:** Pole `UpdateAction` v `config.json` je volitelné a přepisuje globální parametr `-UpdateAction` pro danou konkrétní úlohu. Pokud není uvedeno, úloha dědí globální hodnotu (výchozí: `RenameNew`).
```

| Pole | Povinné | Popis |
|---|---|---|
| `Type` | Ano | `"SharePoint"` |
| `SiteName` | Ano* | Zobrazovaný název webu. Skript vyhledá odpovídající web a preferuje přesnou shodu s DisplayName. |
| `SiteUrl` | Ano* | Identifikátor webu. Přijímá plné URL (`https://contoso.sharepoint.com/sites/Marketing`), krátký formát (`contoso.sharepoint.com:/sites/Marketing`) i bez protokolu (`contoso.sharepoint.com/sites/Marketing`). Všechny formáty se automaticky převedou na formát Graph API. Použijte místo `SiteName` pro zamezení nejednoznačných výsledků vyhledávání. |
| `LibraryName` | Ano | Název knihovny dokumentů. Běžné hodnoty: `"Shared Documents"`, `"Documents"`. Skript řeší aliasy mezi těmito názvy a také zkouší porovnání s URL-dekódovanou adresou. |
| `SourcePath` | Ne | Cesta k podsložce v rámci knihovny pro zálohu. Použijte `"/"` nebo vynechte pro kořen knihovny. **Nezahrnujte** název knihovny do této cesty. |
| `LocalDownloadPath` | Ano | Lokální adresář pro uložení souborů. Vytvoří se automaticky, pokud neexistuje. |
| `UpdateAction` | Ne | Přepsání globálního parametru `-UpdateAction` pro danou úlohu. `"Overwrite"` přepíše existující soubory. `"RenameNew"` přejmenuje existující soubor s příponou `_prev_XXXXX` a stáhne novou verzi pod původním názvem. Pokud není uvedeno, dědí globální hodnotu (výchozí: `RenameNew`). |

\* Uveďte buď `SiteName`, nebo `SiteUrl`. `SiteUrl` je doporučeno pro spolehlivost.

### Zálohování dokumentů z kanálů Teams

Teams ukládá soubory kanálů na SharePoint webu. Použijte typ `"SharePoint"` a nasměrujte ho na web daného týmu:

```json
{
  "Type": "SharePoint",
  "SiteUrl": "contoso.sharepoint.com:/sites/SalesTeam",
  "LibraryName": "Shared Documents",
  "SourcePath": "/General",
  "LocalDownloadPath": "C:\\Backups\\SalesTeam-General"
}
```

- Každý kanál Teams má složku uvnitř `Shared Documents` pojmenovanou podle kanálu (např. `/General`, `/Project Alpha`).
- Pro zálohu všech kanálů najednou nastavte `SourcePath` na `"/"`.
- Pro zjištění SharePoint URL vašeho týmu: otevřete kanál v Teams, klikněte na **Otevřít v SharePointu** na kartě Soubory a poznamenejte si URL webu.

### Úloha OneDrive

```json
{
  "Type": "OneDrive",
  "TargetUser": "jana.novakova@contoso.com",
  "SourcePath": "/",
  "LocalDownloadPath": "C:\\Backups\\JanaNovakova-OneDrive"
}
```

| Pole | Povinné | Popis |
|---|---|---|
| `Type` | Ano | `"OneDrive"` |
| `TargetUser` | Ne | UPN uživatele, jehož OneDrive se má zálohovat (např. `jana.novakova@contoso.com`). Pokud není uvedeno, skript použije OneDrive přihlášeného uživatele. Povinné pro scénáře app-only nebo správcovského přístupu. |
| `SourcePath` | Ne | Cesta k podsložce v rámci OneDrive. `"/"` nebo vynechte pro kořen. |
| `LocalDownloadPath` | Ano | Lokální adresář pro uložení souborů. |
| `UpdateAction` | Ne | Stejné jako u SharePoint úloh. |

### Příklad s více úlohami

```json
[
  {
    "Type": "SharePoint",
    "SiteUrl": "contoso.sharepoint.com:/sites/Engineering",
    "LibraryName": "Shared Documents",
    "SourcePath": "/",
    "LocalDownloadPath": "C:\\Backups\\Engineering"
  },
  {
    "Type": "SharePoint",
    "SiteName": "HR Portal",
    "LibraryName": "Policies",
    "SourcePath": "/2024",
    "LocalDownloadPath": "C:\\Backups\\HR-Policies-2024"
  },
  {
    "Type": "OneDrive",
    "TargetUser": "admin@contoso.com",
    "SourcePath": "/Projects",
    "LocalDownloadPath": "C:\\Backups\\Admin-Projects"
  }
]
```

## Logika inkrementálního zálohování

Skript přeskakuje soubory, které se od poslední zálohy nezměnily:

1. **SHA256 / SHA1 hash** — Použije se, když vzdálený soubor poskytuje SHA hash (typické pro osobní OneDrive). Vypočítá se hash lokálního souboru a porovná se.
2. **Velikost + datum poslední změny** — Použije se, když je k dispozici pouze `quickXorHash` (typické pro SharePoint a Teams soubory). Porovnává velikost souboru a `lastModifiedDateTime` s tolerancí 2 sekundy.
3. **Pouze velikost** — Poslední možnost, když z API není k dispozici žádný hash ani datum.

Po stažení souboru skript nastaví lokální časové razítko poslední změny souboru na hodnotu ze vzdáleného zdroje, aby porovnání na základě data správně fungovalo při dalších spuštěních.

## Výstupní soubory

Všechny výstupní soubory se vytvářejí v adresáři skriptu:

| Soubor | Popis |
|---|---|
| `script_log_YYYYMMDD-HHmmss.txt` | Kompletní log pro každé spuštění. Při každém spuštění se vytvoří nový soubor. |
| `renamed_files_manifest.csv` | Evidence souborů přejmenovaných kvůli neplatným znakům v původním názvu. Mapuje původní název na bezpečný název s ID položky pro dohledatelnost. |

## Zkušební běh (Dry Run)

Použijte `-DryRun` pro náhled zálohy bez zápisu jakýchkoli souborů:

```powershell
.\Backup-M365-Interactive.ps1 -DryRun
.\Backup-M365-Interactive.ps1 -DryRun -UpdateAction Overwrite
```

Zaloguje, co by se stáhlo, přeskočilo nebo aktualizovalo, aniž by provedl jakékoli změny na disku.

## GUI

Spusťte `Backup-GUI.ps1` pro grafické rozhraní. Na nástrojové liště je rozbalovací nabídka **Update Mode** (`RenameNew` / `Overwrite`), která nastavuje globální parametr `-UpdateAction` pro proces zálohy. Výchozí výběr je `RenameNew`.

## Omezování a zpracování chyb

- **Odpovědi 429 / 503 / 504** se automaticky opakují s exponenciálním odstupem (až 10 pokusů). Hlavička `Retry-After` je respektována, pokud je přítomna.
- **Ostatní chyby** (404, 401 atd.) se neopakují. Úloha zaloguje chybu a pokračuje dalším souborem nebo úlohou.
- **Sešity OneNote** a další položky typu balíček (package) v knihovnách dokumentů jsou detekovány a přeskočeny s varováním (nelze je stáhnout jako běžné soubory přes Graph API).

## Řešení problémů

### "Library not found" na webu Teams

Log zobrazí všechny dostupné disky na daném webu. Zkontrolujte přesnou hodnotu `Name` a použijte ji jako `LibraryName` v konfiguraci. U většiny Teams webů se knihovna jmenuje `"Shared Documents"` nebo `"Documents"`.

### Nalezen nesprávný web

Pokud je `SiteName` nejednoznačný (např. "Sales" odpovídá "Sales", "Pre-Sales", "Sales Reports"), log upozorní a zobrazí, který web byl vybrán. Přepněte na `SiteUrl` pro přesné vyhledání:

```json
"SiteUrl": "contoso.sharepoint.com:/sites/Sales"
```

### Chyby validace konfigurace

Chybějící nebo neplatná pole jsou hlášena na začátku každé úlohy se srozumitelnou zprávou. Úloha se přeskočí a skript pokračuje zbývajícími úlohami.

### Jak zjistit hodnotu SiteUrl

1. Otevřete SharePoint web nebo kartu **Soubory** kanálu Teams v prohlížeči.
2. Zkopírujte URL z adresního řádku — můžete ji použít přímo:

```json
"SiteUrl": "https://contoso.sharepoint.com/sites/MujWeb"
```

Skript automaticky převede jakýkoli z těchto formátů na formát Graph API:

| Vstupní formát | Příklad |
|---|---|
| Plné URL | `https://contoso.sharepoint.com/sites/Sales` |
| Bez protokolu | `contoso.sharepoint.com/sites/Sales` |
| Graph formát (s dvojtečkou) | `contoso.sharepoint.com:/sites/Sales` |

Všechny tři formáty jsou ekvivalentní. Skript je interně převede na `contoso.sharepoint.com:/sites/Sales` (formát vyžadovaný Microsoft Graph API).

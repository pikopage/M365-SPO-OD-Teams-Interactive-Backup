# Microsoft Graph - Delegated vs Application Permissions

## Dva typy oprávnění

| | Delegated (delegovaná) | Application (aplikační) |
|---|---|---|
| **Kdo se přihlašuje** | Reálný uživatel (popup v prohlížeči) | Nikdo — aplikace se autentizuje certifikátem nebo secretem |
| **Rozsah přístupu** | Pouze to, k čemu má přihlášený uživatel přístup | Vše v celém tenantovi |
| **`Files.Read.All` znamená** | "Číst všechny soubory, ke kterým má **tento uživatel** přístup" | "Číst všechny soubory **všech uživatelů** v tenantovi" |
| **Jak se aktivuje v `Connect-MgGraph`** | Parametr `-Scopes` (interaktivní přihlášení) | `-ClientId` + `-CertificateThumbprint` nebo `-ClientSecretCredential` |

Přípona `.All` je zavádějící — u delegovaných oprávnění znamená "všechny soubory, kam se uživatel dostane", ne "všechny soubory v organizaci."

Efektivní přístup je vždy průnik:

```
Efektivní přístup = Udělená oprávnění aplikace  ∩  Vlastní oprávnění uživatele
```

I když má aplikace `Files.Read.All` jako delegované oprávnění, uživatel s přístupem pouze ke 3 SharePoint webům může číst soubory jen z těchto 3 webů.

## Co dělá admin consent (souhlas správce)

Když se při prvním spuštění skriptu zobrazí v prohlížeči prompt **"Consent on behalf of your organization"**, znamená to:

| | Bez admin consentu | S admin consentem |
|---|---|---|
| Typ oprávnění | Delegované | Stále delegované |
| Uživatel se musí přihlásit | Ano | Ano |
| Přístup omezen na data uživatele | Ano | Ano |
| Každý uživatel dostane consent prompt | Ano | **Ne** — admin už schválil za všechny |

Admin consent je čistě administrativní usnadnění — schválit jednou místo toho, aby každý uživatel schvaloval individuálně. **Nepřepíná oprávnění na aplikační úroveň.**

## Jak rozpoznat delegovaná vs aplikační oprávnění

### V kódu (Connect-MgGraph)

```powershell
# DELEGOVANÉ — otevře se prohlížeč, uživatel se přihlásí
Connect-MgGraph -Scopes "Files.Read.All", "Sites.Read.All"

# APLIKAČNÍ — žádný prohlížeč, žádný uživatel, autentizace certifikátem
Connect-MgGraph -ClientId "app-guid" -TenantId "tenant-guid" -CertificateThumbprint "ABC123"
```

Jednoduché pravidlo: pokud se otevře prohlížeč a žádá přihlášení — je to delegované.

### V Azure portálu (Entra ID)

1. Azure Portal > Entra ID > App registrations > najít aplikaci "Microsoft Graph PowerShell"
2. API permissions
3. Zkontrolovat, že všechna oprávnění jsou typu **Delegated**:

   | Oprávnění | Typ | Stav |
   |---|---|---|
   | Files.Read.All | **Delegated** | Granted |
   | Sites.Read.All | **Delegated** | Granted |
   | User.Read | **Delegated** | Granted |

4. Pokud existují oprávnění typu **Application** — odstranit je

## Registrace aplikace v Entra ID — vlastní vs. výchozí

### Výchozí aplikace Microsoft Graph PowerShell

Při spuštění `Connect-MgGraph` bez parametru `-ClientId` se skript přihlásí pod **předregistrovanou aplikací Microsoftu** s názvem *Microsoft Graph PowerShell*:

| Parametr | Hodnota |
|---|---|
| Client ID | `14d82eec-204b-4c2f-b7e8-296a70dab67e` |
| Vlastník | Microsoft (přítomna automaticky v každém tenantu) |
| Typ oprávnění | Delegated — uživatel se musí přihlásit |

Tato aplikace je dostupná každému uživateli bez jakéhokoli nastavení ze strany administrátora. Právě ji tento zálohovací skript používá.

### Vlastní registrace aplikace

Správce může v Entra ID vytvořit vlastní App Registration a přidělit jí přesně definovaná oprávnění:

1. **Entra ID > App registrations > New registration**
2. Přidat **Delegated** oprávnění: `Files.Read.All`, `Sites.Read.All`, `User.Read`
3. Kliknout na **Grant admin consent** — uživatelé pak nedostanou žádný consent prompt
4. Předat uživateli **Application (client) ID** a **Directory (tenant) ID**

V skriptu pak stačí:
```powershell
Connect-MgGraph -ClientId "váš-client-id" -TenantId "váš-tenant-id" `
                -Scopes "Files.Read.All", "Sites.Read.All", "User.Read"
```

| | Výchozí app (Microsoft) | Vlastní App Registration |
|---|---|---|
| Potřeba správce pro nastavení | Ne | Ano (jednorázově) |
| Consent prompt pro uživatele | Ano (poprvé) | Ne (admin pre-consent) |
| Kontrola nad oprávněními | Žádná | Plná |
| Vhodné pro distribuci / automatizaci | Méně | Doporučeno |

---

## WAM — Windows Web Account Manager a tiché přihlášení

### Co je WAM

WAM je systémová komponenta Windows 10/11, která umí vydávat tokeny pro přihlášené uživatele **bez zobrazení jakéhokoli okna**. Funguje pouze na strojích **připojených k Azure AD** (AAD joined nebo Hybrid joined).

### Jak WAM ovlivňuje chování Connect-MgGraph

```
Stroj připojen k Azure AD (firemní počítač)
        │
        └─ Windows udržuje tzv. Primary Refresh Token (PRT)
                │
                └─ MSAL/Connect-MgGraph detekuje PRT přes WAM
                        │
                        └─ Token vydán tiše — žádný popup, žádný prohlížeč ✓

Stroj NENÍ připojen k Azure AD (testovací / domácí počítač)
        │
        └─ PRT neexistuje, WAM nedostupný pro M365
                │
                └─ MSAL musí použít interaktivní přihlášení
                        │
                        └─ Otevře se prohlížeč nebo popup okno ✗ (musí uživatel)
```

### Proč se chování liší mezi testovacím a produkčním prostředím

| Prostředí | Připojení k AAD | WAM dostupný | Výsledek |
|---|---|---|---|
| Produkce (firemní PC) | Ano | Ano | Tiché přihlášení, žádný popup |
| Test (vývojářský / domácí PC) | Ne | Ne | Popup s přihlášením vždy |

**Důležité:** Toto chování není chyba — jde o záměrný bezpečnostní mechanismus Windows. Uživatel na AAD-joined stroji je automaticky identifikován přes své Windows přihlášení.

### Ověření stavu AAD join na stroji

```powershell
dsregcmd /status | Select-String "AzureAdJoined|AzureAdPrt"
```

Výstup na firemním (produkčním) stroji:
```
AzureAdJoined  : YES
AzureAdPrt     : YES     ← WAM bude fungovat tiše
```

Výstup na testovacím / nepřipojeném stroji:
```
AzureAdJoined  : NO
AzureAdPrt     : NO      ← popup se zobrazí
```

---

## Další možnosti omezení přístupu

- **Conditional Access** — omezit, kteří uživatelé mohou aplikaci "Microsoft Graph PowerShell" používat, nebo vyžadovat MFA
- **Admin consent workflow** — vyžadovat schválení správce předtím, než jakýkoli uživatel může udělit souhlas (Entra ID > Enterprise applications > Consent and permissions > User consent settings)
- **Omezení na konkrétní uživatele** — na Enterprise aplikaci nastavit "Assignment required" na Yes, pak přiřadit pouze konkrétní uživatele/skupiny

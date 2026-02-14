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

## Další možnosti omezení přístupu

- **Conditional Access** — omezit, kteří uživatelé mohou aplikaci "Microsoft Graph PowerShell" používat, nebo vyžadovat MFA
- **Admin consent workflow** — vyžadovat schválení správce předtím, než jakýkoli uživatel může udělit souhlas (Entra ID > Enterprise applications > Consent and permissions > User consent settings)
- **Omezení na konkrétní uživatele** — na Enterprise aplikaci nastavit "Assignment required" na Yes, pak přiřadit pouze konkrétní uživatele/skupiny

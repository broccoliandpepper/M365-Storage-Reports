# Microsoft 365 Storage Reporting Scripts

PowerShell scripts to export storage and usage information for Exchange Online mailboxes, OneDrive for Business, SharePoint / Teams sites, and to generate a consolidated Microsoft 365 storage report. [file:3][file:4][file:5][file:6]

Scripts PowerShell pour exporter les informations de stockage et d’usage des boîtes aux lettres Exchange Online, des OneDrive for Business, des sites SharePoint / Teams, ainsi qu’un rapport consolidé Microsoft 365. [file:3][file:4][file:5][file:6]

---

## Repository structure / Structure du dépôt

src/
Export-M365StorageReport.ps1
Export-MailboxStorage.ps1
Export-OneDriveStorage.ps1
Export-TeamsStorage.ps1
Exchange_mailbox_infos.ps1
OneDrive_SharePoint_Infos.ps1
docs/
examples/
sample-output-MailboxStorage.csv (optional)
sample-output-OneDriveStorage.csv (optional)
sample-output-TeamsStorage.csv (optional)
sample-output-M365StorageSummary.csv (optional)
LICENSE
README.md

text

---

## Prerequisites / Prérequis

**EN**

- PowerShell 7.3 or later. [file:3][file:4][file:5][file:6]  
- Ability to install modules in the current user scope (or run as administrator for all-users scope). [file:3][file:4][file:5][file:6]  
- Network access to Microsoft 365 endpoints for Exchange Online and Microsoft Graph. [file:3][file:4][file:5][file:6]

**Required PowerShell modules**

- `ExchangeOnlineManagement` (for Exchange Online mailbox data). [file:3][file:4]  
- Microsoft Graph modules, depending on the script:  
  - `Microsoft.Graph.Authentication` [file:3][file:5][file:6]  
  - `Microsoft.Graph.Reports` [file:3][file:5][file:6]  
  - `Microsoft.Graph.Users` [file:3]  
  - `Microsoft.Graph.Sites` [file:3][file:5][file:6]  
  - `Microsoft.Graph` (meta module, used in OneDrive_SharePoint_Infos.ps1). [file:2]

The scripts automatically check for and install required modules in the appropriate scope when possible. [file:3][file:4][file:5][file:6]

**Microsoft 365 permissions (examples)**

- Exchange Online:
  - Exchange Online Administrator or equivalent role to read mailbox statistics and quotas. [file:3][file:4]
- Microsoft Graph:
  - `User.Read.All`, `Reports.Read.All`, `Sites.Read.All` for storage reports. [file:3][file:5][file:6]
  - `Team.ReadBasic.All`, `AppCatalog.Read.All`, `Directory.Read.All` for advanced Teams / apps info in OneDrive_SharePoint_Infos.ps1. [file:2]

---

**FR**

- PowerShell 7.3 ou supérieur. [file:3][file:4][file:5][file:6]  
- Possibilité d’installer des modules dans le scope utilisateur (ou exécuter en administrateur pour le scope AllUsers). [file:3][file:4][file:5][file:6]  
- Accès réseau aux endpoints Microsoft 365 pour Exchange Online et Microsoft Graph. [file:3][file:4][file:5][file:6]

**Modules PowerShell requis**

- `ExchangeOnlineManagement` (données des boîtes aux lettres Exchange Online). [file:3][file:4]  
- Modules Microsoft Graph, selon le script :  
  - `Microsoft.Graph.Authentication` [file:3][file:5][file:6]  
  - `Microsoft.Graph.Reports` [file:3][file:5][file:6]  
  - `Microsoft.Graph.Users` [file:3]  
  - `Microsoft.Graph.Sites` [file:3][file:5][file:6]  
  - `Microsoft.Graph` (meta module, utilisé dans OneDrive_SharePoint_Infos.ps1). [file:2]

Les scripts vérifient et installent automatiquement les modules requis dans le scope approprié lorsque c’est possible. [file:3][file:4][file:5][file:6]

**Droits Microsoft 365 (exemples)**

- Exchange Online :
  - Rôle Administrateur Exchange Online ou équivalent pour lire les statistiques et quotas des boîtes aux lettres. [file:3][file:4]
- Microsoft Graph :
  - `User.Read.All`, `Reports.Read.All`, `Sites.Read.All` pour les rapports de stockage. [file:3][file:5][file:6]  
  - `Team.ReadBasic.All`, `AppCatalog.Read.All`, `Directory.Read.All` pour les infos avancées Teams / apps dans OneDrive_SharePoint_Infos.ps1. [file:2]

---

## Main script: Export-M365StorageReport.ps1

**EN**

Orchestrator script that runs mailbox, OneDrive and Teams storage collection and generates multiple CSV exports and summaries. [file:3]

**Key features**

- PowerShell version check (requires 7.3+). [file:3]  
- Admin privilege detection to choose installation scope for modules. [file:3]  
- Automatic installation / verification of:
  - `ExchangeOnlineManagement`
  - `Microsoft.Graph.Authentication`, `Microsoft.Graph.Reports`, `Microsoft.Graph.Users`, `Microsoft.Graph.Sites` (depending on included services) [file:3]
- Robust logging with a timestamped log file in the output folder. [file:3]  
- Retry logic for Exchange Online and Microsoft Graph connections. [file:3]  
- Consolidation of data from:
  - Exchange Online mailboxes
  - OneDrive for Business
  - Teams / SharePoint sites [file:3]

**Parameters**

- `-OutputFolder <string>` (mandatory): Folder where CSV reports and log files will be saved. [file:3]  
- `-IncludeTeams [switch]` (optional, default: `$true`): Include Teams storage data. [file:3]  
- `-IncludeMailboxes [switch]` (optional, default: `$true`): Include mailbox storage data. [file:3]  
- `-IncludeOneDrive [switch]` (optional, default: `$true`): Include OneDrive storage data. [file:3]  
- `-MaxRetries <int>` (optional, default: `3`): Maximum retry attempts for failed operations. [file:3]

**Generated files (examples)**

- `MailboxStorageData_YYYYMMDDHHmmss.csv` [file:3]  
- `OneDriveStorageData_YYYYMMDDHHmmss.csv` [file:3]  
- `TeamsStorageData_YYYYMMDDHHmmss.csv` [file:3]  
- `ConsumedDataStorage_YYYYMMDDHHmmss.csv` (consolidated) [file:3]  
- `StorageSummary_YYYYMMDDHHmmss.csv` (summary KPI) [file:3]  
- `ExecutionSummary.txt` [file:3]  
- `M365StorageReport_YYYYMMDDHHmmss.log` [file:3]

**Examples**

Consolidated report with all services
.\src\Export-M365StorageReport.ps1 -OutputFolder "C:\Reports"

Only mailboxes and OneDrive, no Teams
.\src\Export-M365StorageReport.ps1 -OutputFolder "C:\Reports" -IncludeTeams:$false

text

---

**FR**

Script orchestrateur qui lance la collecte de stockage pour les boîtes aux lettres, OneDrive et Teams, puis génère plusieurs exports CSV et des fichiers de synthèse. [file:3]

**Fonctionnalités principales**

- Vérification de la version de PowerShell (nécessite 7.3+). [file:3]  
- Détection des privilèges administrateur pour choisir le scope d’installation des modules. [file:3]  
- Installation / vérification automatique de :
  - `ExchangeOnlineManagement`
  - `Microsoft.Graph.Authentication`, `Microsoft.Graph.Reports`, `Microsoft.Graph.Users`, `Microsoft.Graph.Sites` (selon les services inclus) [file:3]
- Journalisation détaillée via un fichier log horodaté dans le dossier de sortie. [file:3]  
- Mécanismes de retry pour les connexions Exchange Online et Microsoft Graph. [file:3]  
- Consolidation des données provenant de :
  - Boîtes aux lettres Exchange Online
  - OneDrive for Business
  - Sites Teams / SharePoint [file:3]

**Paramètres**

- `-OutputFolder <string>` (obligatoire) : Dossier dans lequel les rapports CSV et les logs seront générés. [file:3]  
- `-IncludeTeams [switch]` (optionnel, défaut : `$true`) : Inclure les données de stockage Teams. [file:3]  
- `-IncludeMailboxes [switch]` (optionnel, défaut : `$true`) : Inclure les données de stockage des boîtes aux lettres. [file:3]  
- `-IncludeOneDrive [switch]` (optionnel, défaut : `$true`) : Inclure les données de stockage OneDrive. [file:3]  
- `-MaxRetries <int>` (optionnel, défaut : `3`) : Nombre maximal de tentatives en cas d’erreurs. [file:3]

**Fichiers générés (exemples)**

- `MailboxStorageData_YYYYMMDDHHmmss.csv` [file:3]  
- `OneDriveStorageData_YYYYMMDDHHmmss.csv` [file:3]  
- `TeamsStorageData_YYYYMMDDHHmmss.csv` [file:3]  
- `ConsumedDataStorage_YYYYMMDDHHmmss.csv` (consolidé) [file:3]  
- `StorageSummary_YYYYMMDDHHmmss.csv` (KPI de synthèse) [file:3]  
- `ExecutionSummary.txt` [file:3]  
- `M365StorageReport_YYYYMMDDHHmmss.log` [file:3]

**Exemples**

Rapport consolidé avec tous les services
.\src\Export-M365StorageReport.ps1 -OutputFolder "C:\Reports"

Seulement boîtes aux lettres et OneDrive, sans Teams
.\src\Export-M365StorageReport.ps1 -OutputFolder "C:\Reports" -IncludeTeams:$false

text

---

## Export-MailboxStorage.ps1

**EN**

Exports Exchange Online mailbox storage usage to a CSV file. [file:4]

- Installs and imports `ExchangeOnlineManagement` if needed. [file:4]  
- Connects to Exchange Online and retrieves all mailboxes. [file:4]  
- Collects mailbox statistics (item count, total size). [file:4]  
- Converts size values to MB and calculates a global total in GB. [file:4]

**Parameter**

- `-OutputFolder <string>` (mandatory): Destination folder for the mailbox CSV and log file. [file:4]

**Output**

- `MailboxStorageData_YYYYMMDDHHmmss.csv` with:
  - ServiceType, DisplayName, UserPrincipalName, PrimarySmtpAddress [file:4]
  - RecipientType, StorageUsedMB, StorageQuotaMB, ItemCount, LastLogonTime, Database, AdditionalInfo (e.g. Archive yes/no) [file:4]

---

**FR**

Exporte l’usage de stockage des boîtes aux lettres Exchange Online dans un fichier CSV. [file:4]

- Installe et importe `ExchangeOnlineManagement` si nécessaire. [file:4]  
- Se connecte à Exchange Online et récupère toutes les boîtes aux lettres. [file:4]  
- Récupère les statistiques des boîtes (nombre d’éléments, taille totale). [file:4]  
- Convertit les tailles en Mo et calcule un total global en Go. [file:4]

**Paramètre**

- `-OutputFolder <string>` (obligatoire) : Dossier de destination pour le CSV des boîtes aux lettres et le fichier log. [file:4]

**Sortie**

- `MailboxStorageData_YYYYMMDDHHmmss.csv` contenant :
  - ServiceType, DisplayName, UserPrincipalName, PrimarySmtpAddress [file:4]  
  - RecipientType, StorageUsedMB, StorageQuotaMB, ItemCount, LastLogonTime, Database, AdditionalInfo (ex. Archive Yes/No) [file:4]

---

## Export-OneDriveStorage.ps1

**EN**

Exports OneDrive for Business storage data to CSV using Microsoft Graph reports. [file:6]

- Installs and imports `Microsoft.Graph.Authentication` and `Microsoft.Graph.Reports`. [file:6]  
- Connects to Microsoft Graph with reporting scopes. [file:6]  
- Retrieves OneDrive usage account details (`Get-MgReportOneDriveUsageAccountDetail`) for the last 30 days (period `D30`). [file:6]  
- Calculates used and allocated storage in MB, plus file counts and last activity date. [file:6]

**Parameter**

- `-OutputFolder <string>` (mandatory): Destination folder for the OneDrive CSV and log file. [file:6]

**Output**

- `OneDriveStorageData_YYYYMMDDHHmmss.csv` with:
  - ServiceType, DisplayName, UserPrincipalName, PrimarySmtpAddress [file:6]
  - StorageUsedMB, StorageQuotaMB, ItemCount, LastLogonTime, AdditionalInfo (e.g. Active Files count) [file:6]

---

**FR**

Exporte les données de stockage OneDrive for Business vers un fichier CSV en s’appuyant sur les rapports Microsoft Graph. [file:6]

- Installe et importe `Microsoft.Graph.Authentication` et `Microsoft.Graph.Reports`. [file:6]  
- Se connecte à Microsoft Graph avec les scopes de reporting. [file:6]  
- Récupère les détails d’usage OneDrive (`Get-MgReportOneDriveUsageAccountDetail`) pour les 30 derniers jours (période `D30`). [file:6]  
- Calcule le stockage utilisé / alloué en Mo, ainsi que les nombres de fichiers et la dernière activité. [file:6]

**Paramètre**

- `-OutputFolder <string>` (obligatoire) : Dossier de destination pour le CSV OneDrive et le fichier log. [file:6]

**Sortie**

- `OneDriveStorageData_YYYYMMDDHHmmss.csv` contenant :
  - ServiceType, DisplayName, UserPrincipalName, PrimarySmtpAddress [file:6]  
  - StorageUsedMB, StorageQuotaMB, ItemCount, LastLogonTime, AdditionalInfo (par ex. Active Files) [file:6]

---

## Export-TeamsStorage.ps1

**EN**

Exports Microsoft Teams storage usage to CSV, based on SharePoint site usage reports from Microsoft Graph. [file:5]

- Installs and imports `Microsoft.Graph.Authentication`, `Microsoft.Graph.Reports`, `Microsoft.Graph.Sites`. [file:5]  
- Clears Graph token cache before connecting to reduce auth issues. [file:5]  
- Connects to Microsoft Graph with `User.Read.All`, `Reports.Read.All`, `Sites.Read.All`. [file:5]  
- Retrieves SharePoint site usage details and filters sites associated with Teams. [file:5]  
- Computes storage used / allocated in MB, file count and last activity date for each Team. [file:5]

**Parameter**

- `-OutputFolder <string>` (mandatory): Destination folder for the Teams CSV and log file. [file:5]

**Output**

- `TeamsStorageData_YYYYMMDDHHmmss.csv` with:
  - ServiceType, DisplayName (Team name), RecipientType (TeamsChannels) [file:5]
  - StorageUsedMB, StorageQuotaMB, ItemCount, LastLogonTime, Database (SharePointTeams), AdditionalInfo (site URL) [file:5]

---

**FR**

Exporte l’usage de stockage des équipes Microsoft Teams vers un fichier CSV, à partir des rapports d’usage des sites SharePoint via Microsoft Graph. [file:5]

- Installe et importe `Microsoft.Graph.Authentication`, `Microsoft.Graph.Reports`, `Microsoft.Graph.Sites`. [file:5]  
- Vide le cache de jetons Graph avant la connexion pour limiter les problèmes d’authentification. [file:5]  
- Se connecte à Microsoft Graph avec `User.Read.All`, `Reports.Read.All`, `Sites.Read.All`. [file:5]  
- Récupère le détail d’usage des sites SharePoint et filtre ceux liés à Teams. [file:5]  
- Calcule le stockage utilisé / alloué en Mo, le nombre de fichiers et la dernière activité pour chaque équipe. [file:5]

**Paramètre**

- `-OutputFolder <string>` (obligatoire) : Dossier de destination pour le CSV Teams et le fichier log. [file:5]

**Sortie**

- `TeamsStorageData_YYYYMMDDHHmmss.csv` contenant :
  - ServiceType, DisplayName (nom de l’équipe), RecipientType (TeamsChannels) [file:5]  
  - StorageUsedMB, StorageQuotaMB, ItemCount, LastLogonTime, Database (SharePointTeams), AdditionalInfo (URL du site) [file:5]

---

## Exchange_mailbox_infos.ps1

**EN**

Script focused on collecting detailed information about Exchange Online mailboxes, similar in spirit to the mailbox storage exporter but more oriented to mailbox information and statistics exploration. [file:1][file:4]

**FR**

Script centré sur la collecte d’informations détaillées sur les boîtes aux lettres Exchange Online, dans l’esprit de l’export de stockage mais davantage orienté exploration des informations et statistiques. [file:1][file:4]

*(You can expand this section once you decide exactly which properties you want to highlight in the README.)* [file:1]

---

## OneDrive_SharePoint_Infos.ps1

**EN**

Interactive script to collect cloud services information from Microsoft Graph. [file:2]

Main capabilities:

- Active OneDrive users and their used size (MB). [file:2]  
- SharePoint sites list (names, URLs, optional last modified date). [file:2]  
- Teams and private channels (team name, channel names, membership type). [file:2]  
- Third-party Teams apps (from the Teams app catalog). [file:2]

The script:

- Checks and installs the `Microsoft.Graph` module. [file:2]  
- Connects to Microsoft Graph with a set of required scopes. [file:2]  
- Validates that the necessary commands are available (`Get-MgUser`, `Get-MgSite`, `Get-MgTeam`, `Get-MgReportOneDriveUsageAccountDetail`, etc.). [file:2]  
- Displays results in a formatted table and exports them to `CloudInfos.csv`. [file:2]

---

**FR**

Script interactif pour collecter des informations de services cloud via Microsoft Graph. [file:2]

Fonctionnalités principales :

- Utilisateurs OneDrive actifs et taille utilisée (Mo). [file:2]  
- Liste des sites SharePoint (noms, URL, date de dernière modification optionnelle). [file:2]  
- Équipes Teams et canaux privés (nom de l’équipe, noms des canaux, type de canal). [file:2]  
- Applications Teams tierces (issues du catalogue d’applications Teams). [file:2]

Le script :

- Vérifie et installe le module `Microsoft.Graph`. [file:2]  
- Se connecte à Microsoft Graph avec un ensemble de scopes requis. [file:2]  
- Vérifie la disponibilité des commandes nécessaires (`Get-MgUser`, `Get-MgSite`, `Get-MgTeam`, `Get-MgReportOneDriveUsageAccountDetail`, etc.). [file:2]  
- Affiche les résultats dans un tableau formaté et les exporte dans `CloudInfos.csv`. [file:2]

---

## Usage

**EN**

1. Clone the repository and open a PowerShell 7.3+ session. [file:3][file:4][file:5][file:6]  
2. Optionally start PowerShell as administrator if you want modules installed in the all-users scope. [file:3][file:4][file:5][file:6]  
3. Run the script you need, for example:

Consolidated storage report
.\src\Export-M365StorageReport.ps1 -OutputFolder "C:\Reports"

Mailboxes only
.\src\Export-MailboxStorage.ps1 -OutputFolder "C:\Reports"

OneDrive only
.\src\Export-OneDriveStorage.ps1 -OutputFolder "C:\Reports"

Teams only
.\src\Export-TeamsStorage.ps1 -OutputFolder "C:\Reports"

Cloud services infos (OneDrive / SharePoint / Teams / apps)
.\src\OneDrive_SharePoint_Infos.ps1

text

---

**FR**

1. Clone le dépôt, puis ouvre une session PowerShell 7.3+. [file:3][file:4][file:5][file:6]  
2. Optionnel : lance PowerShell en administrateur si tu veux installer les modules au niveau machine. [file:3][file:4][file:5][file:6]  
3. Exécute le script souhaité, par exemple :

Rapport de stockage consolidé
.\src\Export-M365StorageReport.ps1 -OutputFolder "C:\Reports"

Boîtes aux lettres uniquement
.\src\Export-MailboxStorage.ps1 -OutputFolder "C:\Reports"

OneDrive uniquement
.\src\Export-OneDriveStorage.ps1 -OutputFolder "C:\Reports"

Teams uniquement
.\src\Export-TeamsStorage.ps1 -OutputFolder "C:\Reports"

Infos services cloud (OneDrive / SharePoint / Teams / apps)
.\src\OneDrive_SharePoint_Infos.ps1

text

---

## Known issues / Limitations  
## Problèmes connus / Limitations

**EN**

- Microsoft Graph reports and usage data can be delayed and may not reflect real-time usage. [file:3][file:5][file:6]  
- Some reports require specific admin roles or consent for the application; without them, the scripts may return partial or empty data. [file:2][file:3][file:5][file:6]  
- Very large tenants (many users / sites / teams) may generate large CSV files and require more time to complete. [file:3][file:5][file:6]

**FR**

- Les rapports Microsoft Graph et les données d’usage peuvent être décalés dans le temps et ne pas refléter l’usage en temps réel. [file:3][file:5][file:6]  
- Certains rapports nécessitent des rôles administrateur spécifiques ou un consentement explicite ; sans cela, les scripts peuvent renvoyer des données partielles ou vides. [file:2][file:3][file:5][file:6]  
- Les très grands tenants (beaucoup d’utilisateurs / sites / équipes) peuvent générer des CSV volumineux et augmenter le temps d’exécution. [file:3][file:5][file:6]

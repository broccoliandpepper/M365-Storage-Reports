# M365-Storage-Reports

PowerShell scripts to export storage and usage information for Exchange Online mailboxes, OneDrive for Business, SharePoint / Teams sites, and a consolidated Microsoft 365 storage report.

![License](https://img.shields.io/github/license/broccoliandpepper/M365-Storage-Reports)
![Last Commit](https://img.shields.io/github/last-commit/broccoliandpepper/M365-Storage-Reports)
![Issues](https://img.shields.io/github/issues/broccoliandpepper/M365-Storage-Reports)

---

## 🎯 Overview

This repository provides a set of PowerShell scripts to help Microsoft 365 administrators **monitor and report storage usage** across Exchange Online, OneDrive, SharePoint and Teams.  
The reports can be used for capacity planning, license optimisation, storage governance and cost control. [web:116][web:117][web:118]  
Outputs are generated as CSV files so they can be easily analysed in Excel, Power BI or any other reporting tool. [web:124]

Typical use cases:

- Identify top storage consumers across services  
- Detect inactive or oversized mailboxes / sites  
- Prepare data for archive / retention / cleanup projects  
- Provide regular reports to management or customers

---

## 🧰 Scripts included

According to the repository description, this project focuses on:

- **Exchange Online** mailbox storage usage  
- **OneDrive for Business** storage usage per user  
- **SharePoint / Teams sites** storage usage  
- A **consolidated storage report** combining the above

> Adapter les noms de scripts et chemins ci‑dessous à la structure réelle de ton repo (par ex. `/Exchange-Report.ps1`, `/OneDrive-Report.ps1`, etc.).

Example structure:

- `Exchange-Storage-Report.ps1` – Export mailbox sizes and quotas  
- `OneDrive-Storage-Report.ps1` – Export OneDrive used/allocated storage  
- `SharePoint-Storage-Report.ps1` – Export SharePoint / Teams site storage  
- `M365-Storage-Consolidated-Report.ps1` – Merge outputs into a single CSV

---

## 🛠️ Prerequisites

- Windows with **PowerShell 5.1** or **PowerShell 7+**  
- Permissions to read Exchange Online, OneDrive and SharePoint usage:
  - Exchange Online: **View-Only Recipients** or higher  
  - SharePoint / OneDrive: SharePoint admin / Global admin (selon ton tenant)  
- Required modules (adapter selon ton implémentation réelle) :
  - `ExchangeOnlineManagement`  
  - `Microsoft.Graph` ou `PnP.PowerShell` / `SharePointOnline`  

Before first run, install modules for the current user:

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module Microsoft.Graph -Scope CurrentUser
# ou
Install-Module PnP.PowerShell -Scope CurrentUser

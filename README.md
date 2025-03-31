# GLPI to Outlook Auto Importer

This PowerShell script automates the process of downloading a GLPI service desk `.ics` calendar file and importing its events directly into a named sub-calendar within Microsoft Outlook.

---

## 📦 Features

- Downloads an `.ics` file from a configured GLPI URL with token authentication.
- Imports event(s) into a custom calendar in Outlook (e.g., `AutoImport`).
- Automatically creates required folders if they don't exist.
- Logs basic output and validation results to the console.

---

## 📁 Project Structure
``` bash
.
├─── v1-0-3.ps1                               # Working alpha version
├─── data.psd1                                # Configuration values for varables used in script
├───dependencies
    └─── ics_download.ps1                     # Downloads Ics file from glpi-url specified in data.psd1 
├───InvokeCredentials
    └─── token_auth_direct_download.ps1       # OLD, no longer in use since authentication does not work due to sso setup
    └─── v1-0-0.ps1                           # OLD, original test
└───tmp                                       # Unused files, that are simple there for test purposes
```

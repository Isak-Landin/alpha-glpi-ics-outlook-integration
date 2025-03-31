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

.
├── v1-0-3.ps1
├── data.psd1
├── dependecies
    ├── tokent_auth_direct_download.ps1
├── LICENSE
└── README.md

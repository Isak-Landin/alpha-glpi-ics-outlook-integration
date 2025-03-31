# GLPI to Outlook Auto Importer

This PowerShell script automates the process of downloading a GLPI service desk `.ics` calendar file and importing its events directly into a named sub-calendar within Microsoft Outlook.

This Script is intended as a workaround. A workaround in two aspects, the first of which is if your organisation is using MS SSO or any other external auth-service that in turn does not allow you to use CalDav or WebCal to synchronize your glpi calendar to your Outlook calendar. The second part of the workaround is simply the fact that the issue can be resolved if you have admin or more specific permission in order to implement a smoother solution; such as creating a user utilizing glpi db username and password that can access your or many other's calendars. The second part would inherently result in Webcal or CalDav synchronization being possible.

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

---
## :wrench: Usage
- Firstly, get your own download link address that is tied to your account, you can find it here (glpi-address.xyz/front/planning.php):
![image](https://github.com/user-attachments/assets/7cf26121-069c-4668-9283-48b8643231a4)

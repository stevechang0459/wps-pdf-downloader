# Windows PowerShell PDF Downloader

[![hackmd-github-sync-badge](https://hackmd.io/@stevechang/Windows_PowerShell_PDF_Downloader/badge)](https://hackmd.io/@stevechang/Windows_PowerShell_PDF_Downloader)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
![PowerShell 5.1+](https://img.shields.io/badge/PowerShell-5.1%2B-blue)
![Windows 10/11](https://img.shields.io/badge/Windows-10%2F11-blue)

A lightweight **Windows PowerShell** utility that batch-downloads **all PDF and ZIP links** discovered on a web page.

![Screenshot 2025-12-28 220242](https://hackmd.io/_uploads/B101tnCXWg.png)

![Screenshot 2025-12-28 220310](https://hackmd.io/_uploads/r1ckK207-l.png)

## 🚀 Quick Start

1. Place **both files** in the same folder:
    - `run_pdf_downloader.bat`
    - `wps_pdf_downloader.ps1`
2. **Double-click** `run_pdf_downloader.bat`.
3. Follow the prompts to enter the **URL** and **Folder**.

> **💡 Pro Tip:** When the "Security Warning" appears, simply type **`A`** and press **Enter** to allow page parsing in the current session.

---

## 📥 Usage & Examples

| Field | Requirement | Input Example | Resulting Behavior |
| :--- | :--- | :--- | :--- |
| **Page URL** | **Required** | `https://example.com/docs` | Scans the page for `.pdf` & `.zip`. |
| **Output Folder** | Optional | *(Leave Blank)* | Saves to the **current folder**. |
| | Optional | `MyDocs` | Creates a **subfolder** named `MyDocs`. |
| | Optional | `C:\Downloads` | Saves to the **absolute path**. |

*Note: Environment variables like `%USERPROFILE%\Downloads` are fully supported.*

---

## ✨ Key Features

- **Smart Scraper:** Parses absolute, relative, and protocol-relative URLs.
- **Robust Downloader:** 3 retry attempts with **Exponential Backoff**.
- **High Compatibility:** Works on **Windows PowerShell 5.1** (using `-UseBasicParsing`) and **PowerShell 7+**.
- **Launcher Included:** The `.bat` file handles `Unblock-File` and `Bypass` policy automatically.

---

## 🛠️ Troubleshooting

- **Execution Policy Error:** Ensure you are launching via the `.bat` file.
- **Blocked Files:** If the script fails to start, right-click `wps_pdf_downloader.ps1` > Properties > **Unblock**.
- **Download Fails:** Verify your internet connection and ensure the URL is accessible via a standard browser.

---

## 🔍 Technical Overview

1. **Fetch:** Retrieves HTML content via `Invoke-WebRequest`.
2. **Parse:** Extracts `href` targets using **Regex**.
3. **Filter:** Identifies links ending in `.pdf` or `.zip`.
4. **Transfer:** Downloads files with a built-in error-handling loop.

---

## 📜 Changelog

- **v1.0:** Initial release (Retry logic, absolute path support, WinPS 5.1 compatibility).

## 📄 License

This project is licensed under the [MIT License](LICENSE).
© 2025 Steve Chang

---

###### tags: `PowerShell` `Windows`

# 🔎 Browser Forensic Tool

**Cross-Browser Digital Forensics & Investigation Toolkit**

A powerful cross-browser forensic and triage application designed for cybersecurity professionals, digital forensic analysts, and law enforcement agencies to extract, analyze, visualize, and export browser artifacts securely.

---

## 📌 Overview

The **Browser Forensic Tool** enables investigators to collect and analyze browser data from multiple browsers simultaneously. It ensures safe database handling, advanced visualization, professional reporting, and cross-platform compatibility.

### 🎯 Primary Use Cases

* Digital Forensics Investigations
* Cybersecurity Incident Response
* Browser Artifact Analysis
* Evidence Collection & Documentation

---

## 🚀 Features

### 🌐 Multi-Browser Support

* Google Chrome (All Chromium-based browsers)
* Opera & Opera GX
* Brave
* Microsoft Edge
* Mozilla Firefox

---

### 📂 Data Extraction Capabilities

| Data Type           | Description                                   |
| ------------------- | --------------------------------------------- |
| Browsing History    | URLs, titles, visit timestamps                |
| Downloads           | File names, source URLs, size, timestamps     |
| Cookies             | Domain, name, value, expiry                   |
| Bookmarks           | Folder structure, URLs, timestamps            |
| Passwords (Secrets) | Saved credentials (decrypted where supported) |

---

### 📊 Advanced Capabilities

* Date range filtering
* Global search across datasets
* URL normalization (removes tracking parameters)
* URL shortening with tooltip support
* Cross-platform password decryption
* Encryption type detection
* Interactive visualization dashboard
* Professional export reports (HTML, CSV, PDF)
* Chart export (PNG / SVG / PDF)
* Custom chart styling presets

---

## 🖥 Supported Platforms

* **Windows**
* **macOS**
* **Linux**
* Python **3.7+**

Minimum 4GB RAM (8GB recommended)

---

## 📦 Installation

### Install Required Dependencies

```bash
pip install tkinter
pip install tkcalendar
pip install matplotlib
pip install numpy
pip install Pillow
pip install pycryptodome
pip install reportlab
```

### Optional Dependencies

```bash
pip install pywin32   # Windows password decryption
pip install requests  # Link checking functionality
```

---

## 📁 Browser Database Access

### Chromium-Based Browsers

* `History`
* `Cookies / Network/Cookies`
* `Login Data`
* `Bookmarks`

### Firefox

* `places.sqlite`
* `cookies.sqlite`
* `logins.json`

---

## 🔐 Password Decryption Support

| Platform | Method                     |
| -------- | -------------------------- |
| Windows  | DPAPI / AES-GCM            |
| macOS    | Keychain / AES-GCM         |
| Linux    | AES-CBC (legacy) / AES-GCM |

---

## 📊 Visualization & Graphs

### Chart Types

* Pie Chart (Browser distribution)
* Line Chart (Daily activity timeline)

### Interactive Features

* Hover details (activity breakdown)
* Click-to-view detailed records
* Dynamic X-axis scaling (based on date range)
* Fullscreen mode
* Graph export (PNG/SVG/PDF)

### Chart Themes

* Default
* Dark
* Professional
* Colorful

---

## 📤 Export Options

### 🌐 HTML Report

* Complete forensic report
* Clickable Table of Contents
* Bookmark statistics
* Professional styling
* URL tooltips
* Full URL preservation

### 📄 PDF Report

* A3 Landscape layout
* Executive summary
* Top visited URLs
* Section navigation
* Appendix for full URLs
* Professional formatting

### 📑 CSV Files

* `overview.csv`
* `browsing_history.csv`
* `downloads.csv`
* `cookies.csv`
* `bookmarks.csv`
* `secrets.csv`

---

## ⌨️ Keyboard Shortcuts

| Shortcut         | Function          |
| ---------------- | ----------------- |
| Ctrl + G         | Toggle Graph      |
| Ctrl + C         | Copy selected row |
| Ctrl + Shift + C | Copy selected row |
| F11              | Fullscreen graph  |
| Escape           | Exit fullscreen   |

---

## ⚙️ Technical Architecture

### Safe Database Handling

1. Creates temporary database copies
2. Opens via SQLite
3. Extracts data
4. Deletes temporary files

Prevents browser locking and preserves original evidence.

### URL Processing

* Removes tracking parameters (UTM, GCLID, FBCLID)
* Preserves full URLs for exports
* Shortened display for readability

### Memory Optimization

* Temporary directories
* Efficient TreeView rendering
* Auto data cleanup on browser switch

---

## 📌 Legal Notice

⚠️ **IMPORTANT DISCLAIMER**

This tool:

* Is NOT listed under Government e-Marketplace (GeM)
* Is NOT an empaneled forensic laboratory under Section 79A of the IT Act, 2000
* Is provided for **educational and investigative purposes only**

Users must:

* Maintain proper chain of custody
* Verify outputs independently
* Use write-blockers when applicable
* Consult legal experts before court submission

The provider assumes no liability for misuse or legal consequences.

---

## 🛡 Evidence Handling Best Practices

* Document every action
* Validate results manually
* Cross-reference findings
* Preserve original data integrity
* Follow forensic SOPs

---

## 🧠 Core Helper Functions

| Function                     | Purpose                    |
| ---------------------------- | -------------------------- |
| `get_platform()`             | Detect OS                  |
| `copy_db_file()`             | Safe DB duplication        |
| `chrome_time_to_datetime()`  | Convert Chrome timestamps  |
| `firefox_time_to_datetime()` | Convert Firefox timestamps |
| `decrypt_chrome_password()`  | Password decryption        |
| `normalize_url()`            | Remove tracking parameters |
| `extract_domain()`           | Extract domain from URL    |

---

## 📌 Future Enhancements

* Live forensic acquisition mode
* Hash verification for reports
* Timeline correlation engine
* Multi-user profile support
* Cloud artifact extraction

---

## 👨‍💻 Author

Developed for digital investigation and cybersecurity research purposes.

---

## ⭐ Support

If you found this project useful:

* ⭐ Star this repository
* 🛠 Contribute improvements
* 🐛 Report issues

---

> “Digital evidence must be collected carefully, preserved correctly, and analyzed responsibly.”

---


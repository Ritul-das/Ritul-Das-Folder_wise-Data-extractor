# 🚀 Smart Data Extractor

> **Collect important Windows user data quickly, safely, and intelligently.**

Smart Data Extractor is a lightweight Python application that automatically collects important user data from a Windows computer and saves it to a USB drive or another storage device.

Instead of copying the entire system, it focuses only on valuable user files while skipping unnecessary Windows folders. This makes the extraction process much faster, cleaner, and more storage efficient.

Whether you're creating a backup, migrating to a new computer, performing digital forensics, responding to an incident, or conducting an authorized security assessment, Smart Data Extractor helps automate the process.

---

# ✨ Features

## 📂 Smart File Collection

Automatically collects files from important user folders such as:

- Desktop
- Documents
- Downloads
- Pictures
- Videos
- Music
- OneDrive
- Favorites
- Contacts
- Recent Files
- Camera Roll
- Screenshots

It keeps the original folder structure, making files easy to locate later.

---

## 🌐 Browser Data Collection

Collects important browser data from supported browsers.

### Google Chrome

- Bookmarks
- Browsing History
- Login Database

### Microsoft Edge

- Bookmarks
- Browsing History
- Login Database

---

## 📶 WiFi Profile Collection

Extracts saved WiFi profiles from Windows.

Information collected includes:

- Network Name (SSID)
- Saved Password (when available)
- Profile Information

Administrator privileges are required for this feature.

---

## 📄 Supports 50+ File Types

Automatically copies many common file formats.

### Documents

- PDF
- DOC
- DOCX
- TXT
- XLS
- XLSX
- PPT
- PPTX
- CSV
- RTF
- ODT
- Markdown

### Images

- JPG
- JPEG
- PNG
- GIF
- BMP
- TIFF
- SVG
- WEBP
- PSD
- RAW

### Videos

- MP4
- AVI
- MKV
- MOV
- WMV
- FLV

### Audio

- MP3
- WAV
- FLAC
- AAC
- M4A

### Archives

- ZIP
- RAR
- 7Z
- TAR
- GZ

### Database Files

- JSON
- XML
- SQL
- SQLite
- DB

### Source Code

- Python
- Java
- C
- C++
- HTML
- CSS
- JavaScript
- PHP

---

## 🚫 Intelligent System Folder Exclusion

To save time and storage space, Smart Data Extractor automatically skips unnecessary Windows folders.

Examples include:

- Windows
- Program Files
- Program Files (x86)
- ProgramData
- Recovery
- Boot
- Windows.old
- Recycle Bin
- System Volume Information
- Pagefile
- Hibernation File

Only useful user data is collected.

---

## 💾 Smart Storage Management

The application is optimized for removable USB drives.

Features include:

- Free space detection
- Storage usage monitoring
- Configurable storage limit
- Safety buffer protection
- Large file detection
- FAT32 compatibility

---

## ⚡ Fast Copy Engine

Smart Data Extractor uses an optimized copy process that provides:

- Faster scanning
- Recursive folder search
- Extension filtering
- Metadata preservation
- Reliable file copying

---

## 📊 Live Progress Display

During extraction, the application displays:

- Current folder
- Files found
- Files copied
- Data copied
- Storage usage
- Progress updates

---

## 📑 Automatic Summary Report

After every extraction, a detailed report is generated automatically.

The report includes:

- Computer Name
- Username
- Date and Time
- Execution Time
- Destination Folder
- Total Files Copied
- Total Data Size
- Storage Used
- Extracted Locations
- WiFi Profiles
- Skipped System Folders

---

## 🔒 Administrator Support

The application automatically checks for administrator privileges.

If required, Windows will request permission to restart the application with elevated access.

This allows the program to collect additional system information when authorized.

---

## 📁 Organized Output

All collected data is saved inside a timestamped folder.

Example:

```
Data-Backup-20260702-145300
│
├── User_Desktop
├── User_Documents
├── User_Downloads
├── User_Pictures
├── User_Videos
├── Browser_Data
├── WiFi_Profiles
└── EXTRACTION_SUMMARY.txt
```

Everything is organized automatically, making restoration simple.

---

# ⚙️ Requirements

- Windows 7 or later
- Python 3.6+
- USB Drive
- Administrator privileges (recommended)

Install dependency:

```bash
pip install psutil
```

---

# 🚀 Installation

Clone the repository.

```bash
git clone https://github.com/yourusername/smart-data-extractor.git
```

Go to the project folder.

```bash
cd smart-data-extractor
```

Install the required package.

```bash
pip install psutil
```

Run the application.

```bash
python data_extractor.py
```

---

# 🎯 Use Cases

Smart Data Extractor is useful for:

- Personal Backups
- Windows Reinstallation
- Computer Migration
- Data Recovery
- IT Administration
- Digital Forensics
- Incident Response
- Authorized Security Assessments

---

# 🛡️ Privacy

Smart Data Extractor works completely offline.

- No cloud uploads
- No internet connection required
- No external servers
- Your data stays on your computer and storage device

---

# ⚠️ Disclaimer

This project is intended only for systems that you own or have explicit authorization to access.

You are responsible for ensuring that your use of this software complies with all applicable laws, organizational policies, and privacy requirements.

The developers are not responsible for misuse of this software.

---

# 🤝 Contributing

Contributions are always welcome.

If you find a bug, have an idea for a new feature, or want to improve the project, feel free to open an issue or submit a pull request.

---

# 📜 License

This project is licensed under the MIT License.

---

# ❤️ Support the Project

If you found this project useful, consider giving it a ⭐ on GitHub.

Your support helps improve the project and encourages future development.

Thank you for using **Smart Data Extractor**.

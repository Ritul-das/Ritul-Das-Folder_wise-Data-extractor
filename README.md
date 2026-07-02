# 🚀 Smart Data Extractor (Pentesting & Security Utility)

## 📌 Overview

Smart Data Extractor is a Python-based Windows utility that automates structured data collection from a target system. It is designed for **authorized security testing environments only**, where controlled data extraction is required for analysis, auditing, or forensic evaluation.

The tool intelligently scans user environments, prioritizes important directories, and copies selected data into a structured backup format on a USB drive (optimized for large-capacity drives such as 128GB pendrives).

It strictly avoids system-critical folders to maintain stability and reduce unnecessary data collection overhead.

---

## ⚠️ Legal & Ethical Notice

This tool is strictly intended for:

- ✔ Authorized penetration testing
- ✔ Security auditing
- ✔ Digital forensics (with permission)
- ✔ System migration & analysis (authorized environments only)

❌ Unauthorized use on systems without explicit permission is strictly prohibited.

The user is fully responsible for complying with all applicable laws and organizational policies.

---

## ⚙️ Core Functionality

Smart Data Extractor performs structured extraction using a multi-stage pipeline:

### 🔄 Workflow Logic

```
USB Detection
    ↓
Permission Validation (Admin Check)
    ↓
Optimized Path Mapping
    ↓
System Folder Filtering
    ↓
Target Data Scanning
    ↓
Smart File Filtering
    ↓
Secure Copy Operation
    ↓
Space Management Control (128GB limit)
    ↓
Report Generation
```

---

## ✨ Key Features

### 📂 Intelligent Data Collection
- Extracts user-level data from structured Windows directories
- Targets Desktop, Documents, Downloads, Pictures, Videos, Music, and AppData

---

### 🚫 System-Level Exclusion Engine
Automatically excludes critical system directories such as:

- Windows system folders
- Program Files / Program Files (x86)
- System Volume Information
- Recovery partitions
- Boot and system files

This ensures safe operation without system disruption.

---

### 🌐 Browser Data Extraction
Supports extraction of browser-related artifacts:

- Google Chrome (Bookmarks, History, Login Data)
- Microsoft Edge (Bookmarks, History, Login Data)

---

### 📶 WiFi Profile Extraction
- Extracts saved WiFi profiles using Windows native commands
- Retrieves SSID and stored credentials (requires admin privileges)

---

### 💾 Smart Storage Management
- Designed for 128GB pendrive optimization
- Implements 100GB safe usage limit
- Includes 2GB safety buffer
- Prevents disk overflow and corruption

---

### 📁 File Type Intelligence (50+ Extensions)
Supports structured extraction of:

- Documents (PDF, DOCX, TXT, XLSX, PPTX, etc.)
- Images (JPG, PNG, GIF, SVG, RAW, PSD, etc.)
- Videos (MP4, MKV, AVI, MOV, etc.)
- Audio (MP3, WAV, FLAC, AAC, etc.)
- Archives (ZIP, RAR, 7Z, TAR, etc.)
- Code files (PY, JS, HTML, CSS, C++, JAVA, PHP)
- Data formats (JSON, XML, SQL, DB)

---

### ⚡ Performance Optimization
- Multi-layer directory scanning
- Chunk-based file copying (128KB buffer)
- Thread-safe design foundation
- Large file skip logic (>2GB limit for FAT32)

---

### 📊 Real-Time Monitoring
During execution, the tool displays:

- Active scanning paths
- File discovery progress
- Copy status updates
- Storage utilization
- Completion tracking

---

### 📄 Automated Reporting System
Generates a structured extraction report containing:

- System information (username, machine name)
- Execution timestamp
- Total files copied
- Total data size transferred
- Path-wise extraction summary
- WiFi credentials (if extracted)
- Skipped system directories list

---

## 🧠 Optimization Strategy

The tool follows a strict prioritization model:

### 🔝 Priority Levels

1. **User Personal Data (Highest Priority)**
   - Desktop, Documents, Downloads

2. **Work & Project Data**
   - Drive-based folders (D:, E:, F:)

3. **Application Data**
   - AppData (Roaming / Local)

4. **Browser Artifacts**
   - Chrome / Edge profiles

5. **Optional Media Files**
   - Music, Videos, Pictures

System files are always excluded regardless of category.

---

## 💻 Requirements

- Windows 7/8/10/11
- Python 3.6+
- Administrator privileges (required for full functionality)
- USB storage device (recommended: 128GB)

### Dependency

```bash
pip install psutil
```

---

## 🚀 Usage

```bash
python data_extractor.py
```

### Execution Flow

1. Insert USB drive
2. Run script as Administrator
3. Tool detects removable drive automatically
4. Confirms available storage
5. Starts controlled extraction
6. Generates structured backup + report

---

## 📦 Output Structure

```
Data-Backup-[TIMESTAMP]/
│
├── User_Desktop/
├── User_Documents/
├── User_Downloads/
├── User_Pictures/
├── AppData/
├── Browser_Data/
├── WiFi_Profiles/
└── EXTRACTION-SUMMARY.txt
```

---

## 🔐 Security Design

- No external network communication
- No cloud upload or remote sync
- All processing performed locally
- No modification of original files
- Read-only extraction model

⚠️ WiFi credentials and browser data are sensitive and should be handled securely after extraction.

---

## 🧪 Use Cases (Authorized Only)

- Penetration testing (authorized environments)
- Security audits and assessments
- Digital forensic investigations
- Incident response analysis
- System migration in enterprise environments
- Controlled security research labs

---

## ⚠️ Limitations

- Large files (>2GB) are skipped (FAT32 compatibility)
- Performance depends on USB speed
- Admin privileges required for full extraction
- Some browser data may be locked during active sessions

---

## 📜 License

MIT License

---

## 🧾 Disclaimer

This software is provided for educational, security testing, and authorized use only.

The developer is not responsible for misuse or unauthorized deployment of this tool.

---

## ⭐ Project Status

✔ Active Development  
✔ Optimized for Large Storage Devices  
✔ Security-Focused Design  
✔ Pentesting-Oriented Utility  

---

## 🤝 Contribution

Pull requests are welcome for:

- Performance improvements
- Cross-platform support
- Additional browser modules
- Enhanced reporting system
- GUI version development

---

## 🧠 Final Note

This tool is built for controlled environments where structured data extraction is required for security analysis and system evaluation.

Use responsibly.
```

# 🚀 Smart Data Extractor

> A structured Windows-based data extraction utility designed for **authorized security testing, forensic analysis, and controlled system auditing**.

---

# 📖 Introduction

Smart Data Extractor is built to observe a Windows system in a controlled and disciplined way. It does not blindly copy everything. Instead, it follows a structured logic where only meaningful user-level data is collected, while system-level files are intentionally ignored.

It works quietly in the background of a system, identifies useful user content, and prepares a structured backup inside a USB drive, especially optimized for large storage devices like 128GB pendrives.

---

# ⚙️ Working Philosophy

The tool follows a step-by-step logical flow rather than random scanning.

```
USB Detection
      ↓
Permission Check (Administrator Access)
      ↓
User Path Mapping
      ↓
System Folder Filtering
      ↓
Target Data Scanning
      ↓
Smart File Copying
      ↓
Space Management Control
      ↓
Report Generation
      ↓
Final Structured Backup
```

Each stage ensures that the system remains stable, predictable, and safe during operation.

---

# ✨ Key Features

## 📂 Structured User Data Collection

The tool focuses only on meaningful user-generated data such as:

- Desktop files  
- Documents  
- Downloads  
- Pictures  
- Videos  
- Music  
- OneDrive data  
- Application-related user folders  

It preserves the original structure so that data remains easy to understand after extraction.

---

## 🌐 Browser Data Handling

When executed in authorized environments, the tool can extract browser-related artifacts such as:

- Bookmarks  
- Browsing history  
- Saved login data  

Supported browsers include Chrome and Edge.

---

## 📶 WiFi Profile Extraction

The tool can retrieve stored WiFi profiles from the system, including network names and saved credentials, when administrator permissions are granted.

This feature is intended strictly for authorized environments such as forensic or administrative recovery scenarios.

---

## 🚫 System Awareness (Smart Filtering)

One of the most important parts of the tool is what it avoids.

It automatically skips:

- Windows system directories  
- Program Files and Program Files (x86)  
- System Volume Information  
- Recovery and boot partitions  
- Temporary system files  

This ensures the operating system is never disturbed during execution.

---

## 💾 Smart Storage Control

The tool is optimized for large pendrives (especially 128GB devices). It carefully tracks storage usage and ensures that:

- A safe storage limit is maintained  
- Oversized files are skipped when necessary  
- The USB drive is never overloaded  

This makes the process stable and predictable.

---

## 📁 File Type Intelligence

It supports a wide range of file formats including:

- Documents (PDF, DOCX, TXT, XLSX, PPTX)  
- Images (JPG, PNG, GIF, RAW, PSD, SVG)  
- Videos (MP4, MKV, AVI, MOV)  
- Audio (MP3, WAV, FLAC)  
- Archives (ZIP, RAR, 7Z)  
- Code files (PY, JS, HTML, C++, Java)  
- Data files (JSON, XML, SQL, DB)  

---

## ⚡ Performance Design

The extraction process is optimized for efficiency:

- Chunk-based file copying  
- Recursive scanning  
- File filtering by extension  
- Progress tracking during execution  
- Safe handling of large directories  

---

## 📊 Report Generation

At the end of every execution, the tool generates a detailed summary that includes:

- System name and user details  
- Total files copied  
- Total data size  
- Time taken  
- Folder-wise breakdown  
- WiFi profile information  
- Skipped system directories  

This ensures full transparency of the extraction process.

---

# 🧠 Internal Logic

The system follows a priority-based model:

1. User personal data (highest priority)  
2. Work-related folders  
3. Application data  
4. Browser artifacts  
5. Optional media files  

System files are always excluded regardless of condition.

---

# 🧪 Intended Use

This tool is strictly designed for:

- Authorized penetration testing  
- Digital forensic analysis  
- System auditing  
- Controlled security research  
- IT administration tasks  
- System migration in approved environments  

---

# ⚠️ Ethical & Legal Notice

This software must only be used on systems where explicit permission has been granted.

Unauthorized use on personal or third-party systems is strictly prohibited.

The responsibility of lawful usage lies entirely with the user.

---

# 📦 Output Structure

All extracted data is stored in a structured format:

```
Data-Backup-[Timestamp]/
│
├── User_Desktop/
├── User_Documents/
├── User_Downloads/
├── User_Pictures/
├── Browser_Data/
├── WiFi_Profiles/
└── EXTRACTION-SUMMARY.txt
```

---

# 🔐 Security Behavior

- No internet communication  
- No external data transmission  
- No modification of original files  
- Local-only processing  
- Read-only extraction model  


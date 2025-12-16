import ctypes
import os
import shutil
import subprocess
import re
from datetime import datetime
import getpass
import concurrent.futures
from concurrent.futures import ThreadPoolExecutor
import threading
import sys
import traceback
import json
import psutil
import platform

# =============================================================================
# OPTIMIZED PATHS FOR 128GB PENDRIVE - SKIPS SYSTEM FOLDERS
# =============================================================================
class OptimizedPaths:
    """Optimized paths for 128GB pendrive - SKIPS system folders"""
    
    @staticmethod
    def get_optimized_paths():
        """Get ONLY user data and important paths - SKIPS system folders"""
        username = getpass.getuser()
        optimized_paths = []
        
        # PATHS TO ALWAYS SKIP (SYSTEM FOLDERS)
        skip_paths = [
            'C:\\Windows\\',
            'C:\\Program Files\\',
            'C:\\Program Files (x86)\\',
            'C:\\$Recycle.Bin',
            'C:\\System Volume Information',
            'C:\\ProgramData',
            'C:\\Boot',
            'C:\\Recovery',
            'C:\\Windows.old',
            'C:\\PerfLogs',
            'C:\\swapfile.sys',
            'C:\\hiberfil.sys',
            'C:\\pagefile.sys',
        ]
        
        def should_skip(path):
            """Check if path should be skipped"""
            path_lower = path.lower()
            for skip in skip_paths:
                if path_lower.startswith(skip.lower()):
                    return True
            return False
        
        print("\n🔍 Collecting OPTIMIZED paths (skipping system folders)...")
        
        # ===========================================
        # USER DATA PATHS - HIGH PRIORITY
        # ===========================================
        user_data_paths = [
            # Niccolo's specific paths - TOP PRIORITY
            ('D_Desktop_Backup', 'D:\\Desktop Backup 12-22'),
            ('D_Download', 'D:\\Download'),
            ('D_Dept_Library', 'D:\\Dept Library'),
            ('D_NAAC_22', 'D:\\NAAC 22'),
            
            ('E_College', 'E:\\college'),
            ('E_Godde_Souvenir', 'E:\\godde souvenir'),
            ('E_MS', 'E:\\MS'),
            ('E_NAAC_23_28', 'E:\\NAAC 23-28'),
            ('E_SSR_Criteria', 'E:\\SSR criteria 4'),
            
            # User Profile - ESSENTIAL
            ('User_Desktop', os.path.join("C:\\Users", username, "Desktop")),
            ('User_Documents', os.path.join("C:\\Users", username, "Documents")),
            ('User_Downloads', os.path.join("C:\\Users", username, "Downloads")),
            ('User_Pictures', os.path.join("C:\\Users", username, "Pictures")),
            
            # AppData - SELECTIVE (important apps only)
            ('AppData_Roaming', os.path.join("C:\\Users", username, "AppData\\Roaming")),
            ('AppData_Local', os.path.join("C:\\Users", username, "AppData\\Local")),
            
            # OneDrive
            ('OneDrive', os.path.join("C:\\Users", username, "OneDrive")),
        ]
        
        # ===========================================
        # DOCUMENTS AND PROJECTS - MEDIUM PRIORITY
        # ===========================================
        documents_paths = [
            # Recent Documents
            ('Recent_Documents', os.path.join("C:\\Users", username, "AppData\\Roaming\\Microsoft\\Windows\\Recent")),
            
            # Desktop files
            ('Desktop_Files', os.path.join("C:\\Users", username, "Desktop")),
            
            # Screenshots
            ('Screenshots', os.path.join("C:\\Users", username, "Pictures\\Screenshots")),
            
            # Camera Roll
            ('Camera_Roll', os.path.join("C:\\Users", username, "Pictures\\Camera Roll")),
        ]
        
        # ===========================================
        # BROWSER DATA - SELECTIVE
        # ===========================================
        browser_paths = [
            # Chrome - Only important data
            ('Chrome_Bookmarks', os.path.join("C:\\Users", username, "AppData\\Local\\Google\\Chrome\\User Data\\Default\\Bookmarks")),
            ('Chrome_History', os.path.join("C:\\Users", username, "AppData\\Local\\Google\\Chrome\\User Data\\Default\\History")),
            ('Chrome_Passwords', os.path.join("C:\\Users", username, "AppData\\Local\\Google\\Chrome\\User Data\\Default\\Login Data")),
            
            # Edge - Only important data
            ('Edge_Bookmarks', os.path.join("C:\\Users", username, "AppData\\Local\\Microsoft\\Edge\\User Data\\Default\\Bookmarks")),
            ('Edge_History', os.path.join("C:\\Users", username, "AppData\\Local\\Microsoft\\Edge\\User Data\\Default\\History")),
            ('Edge_Passwords', os.path.join("C:\\Users", username, "AppData\\Local\\Microsoft\\Edge\\User Data\\Default\\Login Data")),
        ]
        
        # ===========================================
        # OFFICE AND WORK FILES
        # ===========================================
        office_paths = [
            # Office Templates
            ('Office_Templates', os.path.join("C:\\Users", username, "AppData\\Roaming\\Microsoft\\Templates")),
            
            # Excel/Word recent files
            ('Excel_Startup', os.path.join("C:\\Users", username, "AppData\\Roaming\\Microsoft\\Excel\\XLSTART")),
            ('Word_Startup', os.path.join("C:\\Users", username, "AppData\\Roaming\\Microsoft\\Word\\STARTUP")),
        ]
        
        # ===========================================
        # ADDITIONAL USER FOLDERS (if space permits)
        # ===========================================
        additional_paths = [
            ('User_Music', os.path.join("C:\\Users", username, "Music")),
            ('User_Videos', os.path.join("C:\\Users", username, "Videos")),
            ('User_Contacts', os.path.join("C:\\Users", username, "Contacts")),
            ('User_Favorites', os.path.join("C:\\Users", username, "Favorites")),
        ]
        
        # Combine all categories
        all_categories = [
            user_data_paths,      # Tier 1: Highest priority
            documents_paths,      # Tier 2: Important docs
            browser_paths,        # Tier 3: Browser data
            office_paths,         # Tier 4: Office files
            additional_paths,     # Tier 5: Only if space
        ]
        
        # Add only existing and non-system paths
        for category in all_categories:
            for name, path in category:
                # Check if path exists and should not be skipped
                if (isinstance(path, str) and os.path.exists(path) and 
                    not should_skip(path)):
                    optimized_paths.append((name, path))
                    print(f"✅ Added: {name}")
        
        # Add drives D: and E: root folders (excluding system folders)
        drive_roots = ['D:\\', 'E:\\', 'F:\\', 'G:\\', 'H:\\']
        for drive in drive_roots:
            if os.path.exists(drive):
                # Don't add drive root directly, but check its folders
                try:
                    for folder in os.listdir(drive):
                        full_path = os.path.join(drive, folder)
                        if os.path.isdir(full_path) and not should_skip(full_path):
                            optimized_paths.append((f"{drive[0]}_{folder}", full_path))
                except:
                    pass
        
        # Remove duplicates
        seen = set()
        unique_paths = []
        for name, path in optimized_paths:
            if (name, path) not in seen:
                seen.add((name, path))
                unique_paths.append((name, path))
        
        print(f"\n✅ Total optimized paths: {len(unique_paths)}")
        print("⚠️  System folders (Windows, Program Files, etc.) are SKIPPED")
        
        return unique_paths

# =============================================================================
# SMART COPY MANAGER FOR 128GB LIMIT
# =============================================================================
class SmartCopyManager:
    """Manages copying within 128GB pendrive limit"""
    
    def __init__(self, pendrive_path):
        self.pendrive_path = pendrive_path
        self.total_copied = 0
        self.MAX_SPACE_MB = 110 * 1024  # 110GB in MB (safety buffer)
        self.current_usage_mb = 0
        
    def get_pendrive_free_space(self):
        """Get free space on pendrive in MB"""
        try:
            total, used, free = shutil.disk_usage(self.pendrive_path)
            return free // (1024 * 1024)  # Convert to MB
        except:
            return 0
    
    def can_copy_file(self, file_size_mb):
        """Check if we have space for this file"""
        free_space = self.get_pendrive_free_space()
        # Reserve 2GB safety buffer
        return (self.current_usage_mb + file_size_mb) < (self.MAX_SPACE_MB - 2048)
    
    def update_usage(self, file_size_mb):
        """Update current usage"""
        self.current_usage_mb += file_size_mb
        self.total_copied += file_size_mb
    
    def get_progress(self):
        """Get progress percentage"""
        return min(100, (self.current_usage_mb / self.MAX_SPACE_MB) * 100)

# =============================================================================
# OPTIMIZED COPY FUNCTIONS
# =============================================================================
def copy_file_smart(source, destination, copy_manager):
    """Smart file copy with space checking"""
    try:
        if not os.path.exists(source):
            return False, 0
        
        file_size_mb = os.path.getsize(source) // (1024 * 1024)
        
        # Check if we have space
        if not copy_manager.can_copy_file(file_size_mb):
            return False, file_size_mb
        
        # Check if file is too large (>2GB) - might skip if pendrive is FAT32
        if file_size_mb > 2048:  # 2GB limit for FAT32
            print(f"⚠️  Skipping large file (>2GB): {source}")
            return False, file_size_mb
        
        # Create destination directory
        os.makedirs(os.path.dirname(destination), exist_ok=True)
        
        # Copy file
        with open(source, 'rb') as src, open(destination, 'wb') as dst:
            while True:
                chunk = src.read(131072)  # 128KB chunks
                if not chunk:
                    break
                dst.write(chunk)
        
        # Update usage
        copy_manager.update_usage(file_size_mb)
        
        # Preserve metadata
        shutil.copystat(source, destination)
        
        return True, file_size_mb
        
    except Exception as e:
        print(f"❌ Error copying {source}: {str(e)}")
        return False, 0

def scan_path_smart(source_path, file_extensions, max_files=10000):
    """Smart scanning with limits"""
    try:
        files_found = []
        total_size_mb = 0
        
        # Skip system folders
        skip_keywords = ['windows', 'program files', '$recycle.bin', 'system volume information']
        
        for root, dirs, files in os.walk(source_path):
            # Skip system directories
            dirs[:] = [d for d in dirs if not any(skip in root.lower() for skip in skip_keywords)]
            
            for file in files[:500]:  # Limit files per directory
                # Check file extension
                file_ext = os.path.splitext(file)[1].lower()
                if not file_extensions or file_ext in file_extensions:
                    full_path = os.path.join(root, file)
                    
                    # Skip system/hidden files
                    if any(skip in full_path.lower() for skip in skip_keywords):
                        continue
                    
                    try:
                        file_size = os.path.getsize(full_path)
                        # Skip files larger than 4GB (FAT32 limit issue)
                        if file_size > 4 * 1024 * 1024 * 1024:
                            continue
                            
                        files_found.append(full_path)
                        total_size_mb += file_size // (1024 * 1024)
                        
                        # Limit total files for performance
                        if len(files_found) >= max_files:
                            break
                            
                    except:
                        continue
                
                if len(files_found) >= max_files:
                    break
            
            if len(files_found) >= max_files:
                break
        
        print(f"📊 {os.path.basename(source_path)}: {len(files_found)} files, ~{total_size_mb}MB")
        return files_found
        
    except Exception as e:
        print(f"❌ Scan error for {source_path}: {str(e)}")
        return []

# =============================================================================
# MAIN EXTRACTION FUNCTION
# =============================================================================
def optimized_data_extraction():
    """Optimized extraction for 128GB pendrive"""
    try:
        print("="*80)
        print("🚀 OPTIMIZED DATA EXTRACTION FOR 128GB PENDRIVE")
        print("="*80)
        print("⚠️  SYSTEM FOLDERS ARE SKIPPED: Windows, Program Files, etc.")
        print("="*80)
        
        start_time = datetime.now()
        
        # Find USB drive
        print("\n🔍 Searching for USB drive...")
        usb_drives = []
        import string
        for letter in string.ascii_uppercase:
            drive = f"{letter}:\\"
            if os.path.exists(drive):
                try:
                    drive_type = ctypes.windll.kernel32.GetDriveTypeW(drive)
                    if drive_type == 2:  # USB drive
                        usb_drives.append(drive)
                except:
                    pass
        
        if not usb_drives:
            print("❌ No USB drive found! Insert your 128GB pendrive.")
            input("Press Enter to exit...")
            return
        
        drive_letter = usb_drives[0]
        
        # Check pendrive size
        try:
            total, used, free = shutil.disk_usage(drive_letter)
            free_gb = free // (1024**3)
            total_gb = total // (1024**3)
            
            print(f"📦 Pendrive: {drive_letter}")
            print(f"   Size: {total_gb}GB, Free: {free_gb}GB")
            
            if free_gb < 10:
                print("❌ Not enough space! Need at least 10GB free.")
                input("Press Enter to exit...")
                return
                
            if total_gb < 120:
                print("⚠️  Warning: Pendrive appears smaller than 128GB")
                print("   (this is normal due to formatting)")
            
        except Exception as e:
            print(f"❌ Error checking pendrive: {str(e)}")
            input("Press Enter to exit...")
            return
        
        # Create destination folder
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        destination_folder = os.path.join(drive_letter, f"Data-Backup-{timestamp}")
        
        try:
            os.makedirs(destination_folder, exist_ok=True)
            print(f"✅ Created: {destination_folder}")
        except Exception as e:
            print(f"❌ Error creating folder: {str(e)}")
            input("Press Enter to exit...")
            return
        
        # Get optimized paths
        print("\n📁 Getting optimized paths...")
        custom_paths = OptimizedPaths.get_optimized_paths()
        
        if not custom_paths:
            print("❌ No valid paths found!")
            input("Press Enter to exit...")
            return
        
        print(f"\n✅ Will extract from {len(custom_paths)} locations")
        
        # File extensions to copy
        file_extensions = [
            # Documents
            '.txt', '.doc', '.docx', '.pdf', '.rtf', '.odt',
            '.xls', '.xlsx', '.csv', '.ppt', '.pptx',
            '.md', '.tex', '.epub', '.mobi',
            
            # Images
            '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff',
            '.svg', '.webp', '.psd', '.ai', '.eps', '.raw',
            
            # Videos
            '.mp4', '.avi', '.mkv', '.mov', '.wmv', '.flv',
            
            # Audio
            '.mp3', '.wav', '.flac', '.aac', '.m4a',
            
            # Archives
            '.zip', '.rar', '.7z', '.tar', '.gz',
            
            # Data files
            '.json', '.xml', '.csv', '.sql', '.db', '.sqlite',
            
            # Code files
            '.py', '.java', '.cpp', '.c', '.html', '.css', '.js',
            '.php', '.json', '.xml',
        ]
        
        print(f"📄 File types: {len(file_extensions)} extensions")
        
        # Initialize smart copy manager
        copy_manager = SmartCopyManager(drive_letter)
        free_space_mb = copy_manager.get_pendrive_free_space()
        max_space_mb = min(free_space_mb - 2048, copy_manager.MAX_SPACE_MB)  # Reserve 2GB
        
        print(f"\n📊 Space management:")
        print(f"   Pendrive free: {free_space_mb}MB")
        print(f"   Max to use: {max_space_mb}MB ({max_space_mb//1024}GB)")
        print(f"   Safety buffer: 2GB reserved")
        
        # Confirm
        print("\n" + "="*80)
        print("⚠️  CONFIRM EXTRACTION")
        print("="*80)
        print(f"Pendrive: {drive_letter}")
        print(f"Destination: {destination_folder}")
        print(f"Paths to extract: {len(custom_paths)}")
        print(f"Max space to use: {max_space_mb//1024}GB")
        print("\nSystem folders will be skipped.")
        
        confirm = input("\nType 'YES' to continue or anything else to cancel: ").strip()
        if confirm.upper() != 'YES':
            print("❌ Cancelled.")
            return
        
        # Start extraction
        print("\n" + "="*80)
        print("🚀 STARTING EXTRACTION")
        print("="*80)
        
        total_files_copied = 0
        total_size_copied_mb = 0
        path_summary = []
        
        # Process each path in priority order
        for path_name, path_location in custom_paths:
            # Check if we have space
            if copy_manager.current_usage_mb >= max_space_mb:
                print(f"\n⚠️  Space limit reached! ({copy_manager.current_usage_mb}MB used)")
                print("   Stopping extraction.")
                break
            
            print(f"\n📂 Processing: {path_name}")
            print(f"   Location: {path_location}")
            print(f"   Space used: {copy_manager.current_usage_mb}MB / {max_space_mb}MB")
            
            try:
                # Scan for files
                files_to_copy = scan_path_smart(path_location, file_extensions, max_files=5000)
                
                if not files_to_copy:
                    print(f"   No files found or all skipped")
                    path_summary.append(f"{path_name}: 0 files (none found)")
                    continue
                
                print(f"   Found {len(files_to_copy)} files to copy")
                
                # Copy files
                files_copied = 0
                size_copied_mb = 0
                
                for source_file in files_to_copy:
                    # Check space before each file
                    if copy_manager.current_usage_mb >= max_space_mb:
                        print(f"   ⚠️  Space limit reached for this folder")
                        break
                    
                    # Prepare destination path
                    rel_path = os.path.relpath(os.path.dirname(source_file), path_location)
                    dest_dir = os.path.join(destination_folder, path_name, 
                                           rel_path if rel_path != '.' else '')
                    dest_file = os.path.join(dest_dir, os.path.basename(source_file))
                    
                    # Create directory
                    os.makedirs(dest_dir, exist_ok=True)
                    
                    # Copy file
                    success, file_size_mb = copy_file_smart(source_file, dest_file, copy_manager)
                    
                    if success:
                        files_copied += 1
                        size_copied_mb += file_size_mb
                        
                        # Progress update
                        if files_copied % 50 == 0:
                            progress = copy_manager.get_progress()
                            print(f"   Progress: {files_copied}/{len(files_to_copy)} files, {progress:.1f}% space used")
                
                # Update totals
                total_files_copied += files_copied
                total_size_copied_mb += size_copied_mb
                
                # Add to summary
                summary = f"{path_name}: {files_copied} files ({size_copied_mb}MB)"
                path_summary.append(summary)
                
                print(f"✅ {path_name}: {files_copied} files copied ({size_copied_mb}MB)")
                
            except Exception as e:
                error_msg = f"{path_name}: ERROR - {str(e)}"
                path_summary.append(error_msg)
                print(f"❌ {error_msg}")
        
        # Extract WiFi passwords
        print("\n" + "="*80)
        print("🔍 EXTRACTING ADDITIONAL DATA")
        print("="*80)
        
        wifi_data = []
        try:
            print("\n📶 Extracting WiFi passwords...")
            profiles = []
            data = subprocess.check_output(['netsh', 'wlan', 'show', 'profiles']).decode('utf-8', errors="ignore").split('\n')
            for line in data:
                if "All User Profile" in line:
                    profile_name = line.split(":")[1].strip()
                    try:
                        profile_data = subprocess.check_output(['netsh', 'wlan', 'show', 'profile', profile_name, 'key=clear']).decode('utf-8', errors="ignore").split('\n')
                        for line2 in profile_data:
                            if "Key Content" in line2:
                                password = line2.split(":")[1].strip()
                                profiles.append((profile_name, password))
                    except:
                        profiles.append((profile_name, "ERROR"))
            wifi_data = profiles
            print(f"✅ WiFi: {len(wifi_data)} profiles found")
        except:
            wifi_data = ["WiFi extraction failed"]
            print("❌ WiFi extraction failed")
        
        # Save summary
        print("\n" + "="*80)
        print("💾 SAVING SUMMARY")
        print("="*80)
        
        computer_name = os.environ.get('COMPUTERNAME', 'Unknown')
        username = getpass.getuser()
        execution_time = (datetime.now() - start_time).total_seconds()
        
        try:
            summary_file = os.path.join(destination_folder, f"EXTRACTION-SUMMARY-{timestamp}.txt")
            with open(summary_file, 'w', encoding='utf-8') as f:
                f.write("="*80 + "\n")
                f.write("DATA EXTRACTION SUMMARY - OPTIMIZED FOR 128GB\n")
                f.write("="*80 + "\n\n")
                
                f.write("EXTRACTION DETAILS:\n")
                f.write("-" * 40 + "\n")
                f.write(f"Computer Name: {computer_name}\n")
                f.write(f"Username: {username}\n")
                f.write(f"Extraction Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"Execution Time: {execution_time:.2f} seconds\n")
                f.write(f"Pendrive: {drive_letter}\n")
                f.write(f"Destination: {destination_folder}\n")
                f.write(f"Total Files Copied: {total_files_copied}\n")
                f.write(f"Total Size Copied: {total_size_copied_mb}MB ({total_size_copied_mb/1024:.1f}GB)\n")
                f.write(f"Pendrive Space Used: {copy_manager.current_usage_mb}MB ({copy_manager.current_usage_mb/1024:.1f}GB)\n")
                f.write(f"Pendrive Space Free: {copy_manager.get_pendrive_free_space()}MB\n")
                f.write(f"File Types: {len(file_extensions)} extensions\n\n")
                
                f.write("SYSTEM FOLDERS SKIPPED:\n")
                f.write("-" * 40 + "\n")
                f.write("C:\\Windows\\\n")
                f.write("C:\\Program Files\\\n")
                f.write("C:\\Program Files (x86)\\\n")
                f.write("C:\\$Recycle.Bin\n")
                f.write("C:\\System Volume Information\n")
                f.write("C:\\ProgramData\n")
                f.write("C:\\Boot\n")
                f.write("C:\\Recovery\n")
                f.write("C:\\Windows.old\n")
                f.write("C:\\PerfLogs\n\n")
                
                f.write("PATHS EXTRACTED:\n")
                f.write("-" * 40 + "\n")
                for summary in path_summary:
                    f.write(f"{summary}\n")
                
                f.write("\nWiFi PASSWORDS:\n")
                f.write("-" * 40 + "\n")
                for wifi in wifi_data:
                    if isinstance(wifi, tuple):
                        f.write(f"SSID: {wifi[0]}, Password: {wifi[1]}\n")
                    else:
                        f.write(f"{wifi}\n")
            
            print(f"✅ Summary saved: {summary_file}")
            
        except Exception as e:
            print(f"❌ Error saving summary: {str(e)}")
        
        print("\n" + "="*80)
        print("🎉 EXTRACTION COMPLETE!")
        print("="*80)
        print(f"Total files copied: {total_files_copied}")
        print(f"Total size: {total_size_copied_mb}MB ({total_size_copied_mb/1024:.1f}GB)")
        print(f"Pendrive used: {copy_manager.current_usage_mb}MB ({copy_manager.current_usage_mb/1024:.1f}GB)")
        print(f"Pendrive free: {copy_manager.get_pendrive_free_space()}MB")
        print(f"Time taken: {execution_time/60:.1f} minutes")
        print(f"\n📁 Data saved to: {destination_folder}")
        print("\n⚠️  IMPORTANT NOTES:")
        print("   1. System folders were SKIPPED")
        print("   2. Extraction stopped at 100GB limit")
        print("   3. Files >2GB were skipped (FAT32 limit)")
        print("   4. Safely eject pendrive before removing!")
        print("="*80)
        
        input("\nPress Enter to exit...")
        
    except Exception as e:
        print(f"\n❌ Critical error: {str(e)}")
        traceback.print_exc()
        input("\nPress Enter to exit...")

# =============================================================================
# MAIN EXECUTION
# =============================================================================
if __name__ == "__main__":
    try:
        print("="*80)
        print("🚀 OPTIMIZED DATA EXTRACTION - 128GB PENDRIVE")
        print("="*80)
        print("System folders (Windows, Program Files) are SKIPPED")
        print("="*80)
        
        # Check admin privileges
        if not ctypes.windll.shell32.IsUserAnAdmin():
            print("\n⚠️  Requesting administrator privileges...")
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
            sys.exit(0)
        
        # Start extraction
        optimized_data_extraction()
        
    except Exception as e:
        print(f"\n❌ Critical error: {str(e)}")
        traceback.print_exc()
        input("\nPress Enter to exit...")
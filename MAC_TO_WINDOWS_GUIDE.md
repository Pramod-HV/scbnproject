# Building Windows Installer from macOS

## 🎯 Your Situation

You're on **macOS** and need to create a **Windows installer (.exe)** for Windows users.

---

## ⚠️ The Challenge

**Problem**: Inno Setup (Windows installer creator) only runs on Windows, not macOS.

**Solutions**: You have 3 options:

---

## 🔧 Option 1: Use a Windows Machine or VM (Recommended)

### What You Need:
- Access to a Windows computer, OR
- Windows Virtual Machine on your Mac

### Steps:

#### A. If You Have a Windows Computer:

1. **Transfer your project** to the Windows machine:
   ```bash
   # On Mac: Create a ZIP of your project
   cd /path/to/parent/folder
   zip -r Integrated_Folder.zip Integrated_Folder/
   
   # Transfer via:
   # - USB drive
   # - Email (if under 25MB)
   # - Cloud storage (Google Drive, Dropbox)
   # - Network share
   ```

2. **On Windows machine**:
   ```cmd
   # Extract the ZIP
   # Open Command Prompt in the extracted folder
   
   # Install Python (if not installed)
   # Download from: https://www.python.org/downloads/
   
   # Install Inno Setup
   # Download from: https://jrsoftware.org/isdl.php
   # Run installer, use default settings
   
   # Build the application
   cd Integrated_Folder
   build.bat
   
   # Create the installer
   create_installer.bat
   
   # Result: installer_output/SterlingMapMigrator_Setup_v1.0.0.exe
   ```

3. **Transfer installer back to Mac**:
   - Copy `installer_output/SterlingMapMigrator_Setup_v1.0.0.exe` back to your Mac
   - Upload to cloud storage
   - Distribute to users

#### B. If You Want to Use a Virtual Machine:

**Install Windows VM on Mac:**

1. **Download Parallels Desktop** (paid, easiest):
   - https://www.parallels.com/
   - 14-day free trial available
   - Install Windows 11 (free for development)

2. **Or use VirtualBox** (free):
   - Download VirtualBox: https://www.virtualbox.org/
   - Download Windows 11 ISO: https://www.microsoft.com/software-download/windows11
   - Create new VM, install Windows
   - Allocate at least 4GB RAM, 50GB disk

3. **Then follow steps from Option 1A** inside the VM

---

## 🌐 Option 2: Use GitHub Actions (Automated, Free)

This builds the Windows installer automatically in the cloud!

### Steps:

1. **Push your code to GitHub**:
   ```bash
   cd Integrated_Folder
   git init
   git add .
   git commit -m "Initial commit"
   
   # Create repo on GitHub, then:
   git remote add origin https://github.com/yourusername/your-repo.git
   git push -u origin main
   ```

2. **Create GitHub Actions workflow**:
   
   Create file: `.github/workflows/build-windows.yml`
   
   ```yaml
   name: Build Windows Installer
   
   on:
     push:
       branches: [ main ]
     workflow_dispatch:
   
   jobs:
     build-windows:
       runs-on: windows-latest
       
       steps:
       - uses: actions/checkout@v3
       
       - name: Set up Python
         uses: actions/setup-python@v4
         with:
           python-version: '3.11'
       
       - name: Install dependencies
         run: |
           python -m pip install --upgrade pip
           pip install pyinstaller
           pip install -r requirements.txt
       
       - name: Build application
         run: |
           pyinstaller build_app.spec --clean --noconfirm
       
       - name: Install Inno Setup
         run: |
           choco install innosetup -y
       
       - name: Create installer
         run: |
           & "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer_windows.iss
       
       - name: Upload installer
         uses: actions/upload-artifact@v3
         with:
           name: windows-installer
           path: installer_output/*.exe
   ```

3. **Trigger the build**:
   - Push to GitHub
   - Go to "Actions" tab on GitHub
   - Build will run automatically
   - Download the installer from "Artifacts"

4. **Distribute**:
   - Download the `.exe` from GitHub Actions
   - Upload to your preferred hosting
   - Share with users

---

## 📦 Option 3: Build on Mac, Package Later (Simplest for Testing)

**For testing purposes**, you can skip the installer and just send a ZIP file:

### Steps:

1. **Build on Mac** (but target Windows):
   
   Unfortunately, PyInstaller can't cross-compile. You need Windows to build for Windows.
   
   **However**, you can:
   - Build the macOS version on your Mac
   - Test it thoroughly
   - Then use Option 1 or 2 to create Windows version

2. **Alternative: Send ZIP for now**:
   ```bash
   # Build on Mac
   cd Integrated_Folder
   ./build.sh
   
   # Create macOS installer
   ./create_installer_mac.sh
   
   # For Windows users, tell them:
   # "Windows version coming soon, here's macOS version for Mac users"
   ```

---

## 🎯 Recommended Approach for You

Since you're on Mac and need Windows installer:

### **Best Option: GitHub Actions (Option 2)**

**Why?**
- ✅ Free
- ✅ Automated
- ✅ No Windows machine needed
- ✅ Repeatable
- ✅ Can build for multiple platforms

**Steps Summary:**

1. **Create GitHub repo** (if you don't have one)
2. **Add the workflow file** I provided above
3. **Push your code**
4. **Wait for build** (5-10 minutes)
5. **Download installer** from Actions artifacts
6. **Distribute to users**

### **Alternative: Use a Friend's Windows PC**

If you have access to any Windows computer:
1. Copy your project folder
2. Install Python + Inno Setup
3. Run `build.bat` then `create_installer.bat`
4. Copy the installer back

---

## 📋 Detailed GitHub Actions Setup

### Step 1: Create GitHub Repository

```bash
# On your Mac
cd Integrated_Folder

# Initialize git (if not already)
git init

# Create .gitignore (already exists)
# Add all files
git add .
git commit -m "Add Sterling Map Migrator application"

# Create repo on GitHub.com, then:
git remote add origin https://github.com/YOUR_USERNAME/sterling-map-migrator.git
git branch -M main
git push -u origin main
```

### Step 2: Create Workflow Directory

```bash
mkdir -p .github/workflows
```

### Step 3: Create Workflow File

Create `.github/workflows/build-windows.yml` with the content I provided above.

### Step 4: Push and Build

```bash
git add .github/workflows/build-windows.yml
git commit -m "Add Windows build workflow"
git push
```

### Step 5: Monitor Build

1. Go to your GitHub repository
2. Click "Actions" tab
3. You'll see the workflow running
4. Wait for it to complete (green checkmark)
5. Click on the workflow run
6. Download the artifact (installer)

### Step 6: Distribute

1. Extract the downloaded ZIP
2. You'll have `SterlingMapMigrator_Setup_v1.0.0.exe`
3. Upload to Google Drive, Dropbox, etc.
4. Share with Windows users

---

## 🔄 Workflow for Future Updates

Once set up with GitHub Actions:

```bash
# Make changes to your code
# Commit and push
git add .
git commit -m "Update feature X"
git push

# GitHub automatically builds new installer
# Download from Actions tab
# Distribute new version
```

---

## 💡 Quick Decision Guide

**Choose GitHub Actions if:**
- ✅ You're comfortable with Git/GitHub
- ✅ You want automated builds
- ✅ You'll update the app frequently
- ✅ You don't have Windows access

**Choose Windows VM if:**
- ✅ You want full control
- ✅ You need to test on Windows anyway
- ✅ You have Parallels or VirtualBox
- ✅ You prefer local builds

**Choose Friend's Windows PC if:**
- ✅ Quick one-time build needed
- ✅ You have easy access to Windows
- ✅ Simple and straightforward
- ✅ No setup required

---

## 🆘 Need Help?

### For GitHub Actions:
1. Create the workflow file exactly as shown
2. Push to GitHub
3. Check Actions tab for errors
4. Common issues:
   - File paths (use forward slashes)
   - Missing dependencies (add to requirements.txt)
   - Inno Setup path (usually correct by default)

### For Windows VM:
1. Allocate enough resources (4GB RAM minimum)
2. Install Python first, then Inno Setup
3. Use Command Prompt, not PowerShell
4. Run as Administrator if needed

---

## ✅ Summary

**You're on Mac, need Windows installer:**

1. **Easiest**: Use GitHub Actions (automated, free)
2. **Most Control**: Use Windows VM (Parallels/VirtualBox)
3. **Quickest**: Borrow a Windows PC

**My Recommendation**: Start with GitHub Actions. It's free, automated, and you'll have it set up for future builds.

---

## 🎉 Next Steps

1. Choose your approach (I recommend GitHub Actions)
2. Follow the detailed steps above
3. Build your Windows installer
4. Test it (ask a Windows user to try)
5. Distribute to your users

You've got this! 🚀
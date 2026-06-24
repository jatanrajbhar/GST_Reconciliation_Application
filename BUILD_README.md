# GST Reconciliation Tool - Desktop Application Build Guide

## Overview

This guide explains how to build the GST Reconciliation Tool as a standalone Windows desktop application with an installer.

## Prerequisites

1. **Python 3.8 or higher** - Download from [python.org](https://www.python.org/downloads/)
2. **Inno Setup 6** - Download from [jrsoftware.org](https://jrsoftware.org/isdl.php) (required for creating setup.exe)

## Quick Build (Recommended)

Simply run the build script:

```batch
build.bat
```

This will:
1. Install all required Python dependencies
2. Convert images to required formats
3. Build the executable using PyInstaller
4. Create the installer using Inno Setup

## Manual Build Steps

### Step 1: Install Dependencies

```batch
pip install -r requirements_desktop.txt
```

### Step 2: Run the Build Script

```batch
python build.py
```

## Output Files

After a successful build, you will find:

| Location | Description |
|----------|-------------|
| `dist/GST Reconciliation Tool/` | Portable application (can be run without installation) |
| `installer_output/GST_Reconciliation_Tool_Setup.exe` | Windows installer |

## Running the Desktop Application

### For Development/Testing

```batch
run_desktop.bat
```

Or directly:

```batch
python gst_reconciliation_app.py
```

### Portable Version

Navigate to `dist/GST Reconciliation Tool/` and run:

```
GST Reconciliation Tool.exe
```

### Installed Version

After running the installer, find the application in:
- Start Menu: GST Reconciliation Tool
- Desktop shortcut (if selected during installation)

## Files in This Project

| File | Purpose |
|------|---------|
| `gst_reconciliation_app.py` | Main desktop application (CustomTkinter GUI) |
| `gst_reconciliation.spec` | PyInstaller configuration |
| `installer.iss` | Inno Setup installer script |
| `build.py` | Automated build script |
| `build.bat` | Windows batch file for easy building |
| `run_desktop.bat` | Run the app without building |
| `requirements_desktop.txt` | Python dependencies |
| `logo small.png` | Application icon |
| `Untitled design (2).png` | Installer header image |
| `Privacy Policy.docx` | Privacy policy document |
| `Terms and Conditions.docx` | Terms and conditions document |
| `Template.xlsx` | Sample data template |

## Troubleshooting

### "Python is not installed"

Make sure Python is installed and added to PATH. During Python installation, check "Add Python to PATH".

### "Inno Setup not found"

Download and install Inno Setup 6 from [jrsoftware.org](https://jrsoftware.org/isdl.php). The build script will still create the portable application even without Inno Setup.

### Application won't start

1. Make sure all dependencies are installed: `pip install -r requirements_desktop.txt`
2. Check if antivirus is blocking the application
3. Try running from command line to see error messages

### Images not showing in installer

Run `python build.py` to convert images to the correct formats (ICO and BMP).

## Customization

### Changing the Icon

Replace `logo small.png` with your new logo and run the build script again.

### Changing Installer Text

Edit `installer.iss` to modify:
- Application name
- Publisher information
- Version number
- Installation messages

### Changing Application Behavior

Edit `gst_reconciliation_app.py` to modify the application functionality.

## Support

- Website: https://www.gscintime.com
- Email: info@gscintime.com
- Phone: +91-22-4612 5600

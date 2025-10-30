# Quick Install Guide

## Step-by-Step Installation

### Step 1: Download the Files

1. Go to: https://github.com/aiva-solutions/farah-jml-script
2. Click the green **"Code"** button
3. Select **"Download ZIP"**
4. The file `farah-jml-script-main.zip` will download to your Downloads folder

### Step 2: Extract the Zip File

1. Open your **Downloads** folder
2. Right-click `farah-jml-script-main.zip`
3. Select **"Extract All..."**
4. Choose a location (we recommend `C:\` or your Documents folder)
5. Click **"Extract"**
6. You'll see a folder called `farah-jml-script-main`

**Recommended location:** Extract to `C:\` so the path is `C:\farah-jml-script-main\`

### Step 3: Open PowerShell as Administrator

1. Press the **Windows key** and type: `PowerShell`
2. Right-click on **"Windows PowerShell"** or **"PowerShell"**
3. Select **"Run as Administrator"**
4. Click **"Yes"** if prompted by User Account Control

### Step 4: Navigate to the Extracted Folder

In the PowerShell window, type one of these commands:

**If you extracted to C:\:**
```powershell
cd C:\farah-jml-script-main
```

**If you extracted to Downloads:**
```powershell
cd $env:USERPROFILE\Downloads\farah-jml-script-main
```

**If you extracted to Documents:**
```powershell
cd $env:USERPROFILE\Documents\farah-jml-script-main
```

Press **Enter** after typing the command.

### Step 5: Allow Script Execution (First Time Only)

If this is your first time running PowerShell scripts, type:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```
Press **Enter**, then type `Y` and press **Enter** again.

### Step 6: Install the Macro

Make sure all Word documents are closed, then type:
```powershell
.\Install-Macro.ps1
```

Press **Enter**. You should see:
```
Installing Learner Initials Macro...
Opening Microsoft Word...
Accessing Normal template...
Installing macro code...
✓ Macro installed successfully!
```

### Step 7: Test the Installation

1. Open Microsoft Word
2. Create a new document
3. Type: `Jared Learner` or `JL` (with uppercase J and L)
4. Save and close the document
5. Reopen the document
6. You should see a popup asking: **"Include Mr. Learner's initial?"**

---

## Uninstall

If you need to remove the macro:

1. Open PowerShell as Administrator
2. Navigate to the folder (same as Step 4 above)
3. Type:
```powershell
.\Install-Macro.ps1 -Uninstall
```

Press **Enter**.

---

## Requirements

- **Run PowerShell as Administrator** (required for Normal.dotm access)
- **Close all Word documents** before running the script
- Microsoft Word must be installed

---

## Manual Installation (2 Minutes)

### Step 1: Open VBA Editor
1. Open Word
2. Press `Alt + F11`

### Step 2: Add Code
1. In Project Explorer → **Normal** → **Microsoft Word Objects**
2. Double-click **ThisDocument**
3. Open `learner-initials-macro.bas` in Notepad
4. Copy all code (skip first "Attribute" line if present)
5. Paste into the code window

### Step 3: Save & Test
1. Close VBA Editor (`Alt + Q`)
2. Create a test document with "Jared Learner" or "JL"
3. Save and reopen → Should see popup

---

## Enable Macros

**File** → **Options** → **Trust Center** → **Trust Center Settings** → **Macro Settings**

Choose: **"Disable all macros with notification"** (recommended)

---

## Troubleshooting

**Macro doesn't run?**
- Check macros are enabled (see above)
- Verify code is in `ThisDocument` module
- Restart Word

**Can't save Normal.dotm?**
- Close all Word documents
- Run Word as Administrator

For detailed instructions, see [INSTALLATION.md](INSTALLATION.md)


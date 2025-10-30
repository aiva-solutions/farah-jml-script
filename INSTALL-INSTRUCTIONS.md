# Installation Instructions

Follow these simple steps to install the Learner Initials Macro on your Windows PC.

## Quick Overview

1. Download the ZIP file from GitHub
2. Extract it to a folder on your computer
3. Open PowerShell as Administrator
4. Navigate to the extracted folder
5. Run the installation script

---

## Detailed Step-by-Step Instructions

### Step 1: Download the Files

1. Open your web browser
2. Go to: **https://github.com/aiva-solutions/farah-jml-script**
3. Look for the green **"Code"** button near the top right
4. Click **"Code"** → Select **"Download ZIP"**
5. The file `farah-jml-script-main.zip` will download to your **Downloads** folder

### Step 2: Extract the ZIP File

**Where to extract:** We recommend extracting to `C:\` for the easiest path.

1. Open your **Downloads** folder (usually: `C:\Users\YourName\Downloads`)
2. Find the file: `farah-jml-script-main.zip`
3. **Right-click** on the ZIP file
4. Select **"Extract All..."**
5. In the dialog box:
   - Change the path to: `C:\` (or leave it as default)
   - Check **"Show extracted files when complete"**
   - Click **"Extract"**
6. After extraction, you should see a folder: `C:\farah-jml-script-main`

### Step 3: Open PowerShell as Administrator

1. Press the **Windows key** on your keyboard
2. Type: `powershell`
3. In the search results, you'll see **"Windows PowerShell"**
4. **Right-click** on **"Windows PowerShell"**
5. Select **"Run as Administrator"**
6. A dialog box may appear - click **"Yes"**
7. A blue PowerShell window will open

### Step 4: Navigate to the Extracted Folder

In the PowerShell window that just opened, type this command:

```powershell
cd C:\farah-jml-script-main
```

Press **Enter**.

You should see the prompt change to show: `PS C:\farah-jml-script-main>`

**Note:** If you extracted to a different location, use that path instead:
- For Downloads: `cd $env:USERPROFILE\Downloads\farah-jml-script-main`
- For Documents: `cd $env:USERPROFILE\Documents\farah-jml-script-main`

### Step 5: Allow Script Execution (First Time Only)

If this is the first time you're running a PowerShell script on this computer, type:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

Press **Enter**. Type `Y` and press **Enter** again.

**Note:** This only needs to be done once. You can skip this step if you've already done it before.

### Step 6: Close All Word Documents

**Important:** Make sure Microsoft Word is completely closed before proceeding.

1. Close all Word documents and windows
2. Check the system tray (bottom right) - if you see a Word icon, right-click it and exit
3. Confirm Word is closed

### Step 7: Install the Macro

In PowerShell, type:

```powershell
.\Install-Macro.ps1
```

Press **Enter**.

You should see messages like:
```
Installing Learner Initials Macro...
Opening Microsoft Word...
Accessing Normal template...
Installing macro code...
✓ Macro installed successfully!
```

The installation should complete in a few seconds.

### Step 8: Test the Installation

1. Open Microsoft Word
2. Create a new blank document
3. Type: `Jared Learner` (with uppercase J and L)
4. Save the document (anywhere is fine)
5. **Close the document**
6. **Reopen the same document**
7. You should see a popup asking: **"Include Mr. Learner's initial?"**
8. Click **"Yes"** to test - it should change to "Jared M. Learner"

If you see the popup, the installation was successful! 🎉

---

## Uninstalling the Macro

If you need to remove the macro later:

1. Open PowerShell as Administrator
2. Navigate to the folder:
   ```powershell
   cd C:\farah-jml-script-main
   ```
3. Run:
   ```powershell
   .\Install-Macro.ps1 -Uninstall
   ```

---

## Troubleshooting

### "Cannot find the file"
- Make sure you typed `cd C:\farah-jml-script-main` correctly
- Check that the folder exists at that location

### "Execution policy" error
- Make sure you completed Step 5 (Set-ExecutionPolicy)
- Run it again if needed

### "Access Denied" or "Permission Denied"
- Make sure you opened PowerShell as Administrator (Step 3)
- Close all Word documents and try again

### Macro doesn't work after installation
- Check that macros are enabled in Word: **File** → **Options** → **Trust Center** → **Trust Center Settings** → **Macro Settings** → Select **"Disable all macros with notification"**
- Restart Word and try again

### Still having issues?
See [INSTALLATION.md](INSTALLATION.md) for detailed troubleshooting or manual installation steps.


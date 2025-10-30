# Installation Instructions

## Quick Install

**Prerequisites:**
- Microsoft Word installed on Windows
- Close all Word documents before starting
- No additional PowerShell packages needed (uses built-in features only)

1. **Download:** Go to https://github.com/aiva-solutions/farah-jml-script → Click "Code" → "Download ZIP"
2. **Extract:** Right-click the ZIP → "Extract All" → Extract to `C:\`
3. **Open PowerShell as Administrator:** Windows key → type `powershell` → Right-click → "Run as Administrator"
4. **Run these commands:**
   ```powershell
   cd C:\farah-jml-script-main
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   .\Install-Macro.ps1
   ```
5. **Test:** Create a Word doc with "Jared Learner", save, close, and reopen it

---

## Quick Uninstall

```powershell
cd C:\farah-jml-script-main
.\Install-Macro.ps1 -Uninstall
```

---

## Step-by-Step Details

### Download & Extract

1. Download `farah-jml-script-main.zip` from GitHub
2. Extract to `C:\` (recommended) or your Downloads/Documents folder
3. Result: `C:\farah-jml-script-main\` folder

### Install

1. **Open PowerShell as Administrator:**
   - Press Windows key → Type `powershell` → Right-click → "Run as Administrator" → Click "Yes"

2. **Navigate to folder:**
   ```powershell
   cd C:\farah-jml-script-main
   ```
   *(If extracted elsewhere, use that path instead)*

3. **Allow scripts (first time only):**
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```
   Type `Y` and press Enter

4. **Install:**
   ```powershell
   .\Install-Macro.ps1
   ```
   Wait for "✓ Macro installed successfully!" message

5. **Enable macros in Word:**
   - File → Options → Trust Center → Trust Center Settings → Macro Settings
   - Select "Disable all macros with notification"

### Uninstall

1. Open PowerShell as Administrator
2. Navigate: `cd C:\farah-jml-script-main`
3. Run: `.\Install-Macro.ps1 -Uninstall`

---

## Troubleshooting

**"Execution policy" error:** Run Step 3 above again  
**"Access Denied":** Make sure PowerShell is opened as Administrator  
**Macro doesn't run:** Check macros are enabled in Word (see Install Step 5)  
**File not found:** Verify the folder path is correct

For detailed troubleshooting, see [INSTALLATION.md](INSTALLATION.md)


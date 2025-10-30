# Quick Install Guide

## PowerShell Installation (Automated)

### Install
```powershell
.\Install-Macro.ps1
```

### Uninstall
```powershell
.\Install-Macro.ps1 -Uninstall
```

**Requirements:**
- Run PowerShell as Administrator (for Normal.dotm access)
- Close all Word documents before running
- If prompted, allow script execution: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`

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


# Quick Install Guide

## Install (PowerShell - Recommended)

**Requirements:**
- Microsoft Word installed
- Close all Word documents first
- No PowerShell packages needed (built-in features only)

```powershell
# 1. Download ZIP from GitHub and extract to C:\
# 2. Open PowerShell as Administrator (Windows key → type "powershell" → Right-click → Run as Administrator)

cd C:\farah-jml-script-main
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
.\Install-Macro.ps1
```

## Uninstall

```powershell
cd C:\farah-jml-script-main
.\Install-Macro.ps1 -Uninstall
```

---

## Manual Install (VBA Editor)

1. Open Word → Press `Alt + F11`
2. Project Explorer → **Normal** → **Microsoft Word Objects** → **ThisDocument**
3. Copy code from `learner-initials-macro.bas` → Paste into code window
4. Close VBA Editor (`Alt + Q`)

---

## Enable Macros in Word

**File** → **Options** → **Trust Center** → **Trust Center Settings** → **Macro Settings** → **"Disable all macros with notification"**

---

## Troubleshooting

**Execution policy error:** Run `Set-ExecutionPolicy` command above  
**Access denied:** Open PowerShell as Administrator  
**Macro doesn't run:** Enable macros in Word (see above)

For full instructions, see [INSTALL-INSTRUCTIONS.md](INSTALL-INSTRUCTIONS.md)


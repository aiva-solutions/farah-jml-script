# Installation Guide: Learner Initials Macro

This guide provides step-by-step instructions for installing the VBA macro that automatically adds Jared M. Learner's initials to Word documents.

**Quick Start:** For automated installation, see [QUICK-INSTALL.md](QUICK-INSTALL.md) or run `.\Install-Macro.ps1` in PowerShell.

## Prerequisites

- Microsoft Word installed on Windows
- Administrator access (for Method 1 - Global Template)
- Backup of your Normal.dotm template (recommended)

## Before You Begin: Backup Normal.dotm

**IMPORTANT**: Before modifying Normal.dotm, create a backup:

1. Close Microsoft Word completely
2. Press `Windows + R` to open Run dialog
3. Type: `%APPDATA%\Microsoft\Templates` and press Enter
4. Copy `Normal.dotm` to `Normal.dotm.backup` (or another safe location)

## Method 1: Global Template Installation (Recommended)

This method makes the macro work for **all Word documents** automatically, including:

- New documents created in Word
- Existing documents (.docx, .doc) opened from disk
- Template files (.dotm, .dotx) opened for editing

### Automated Installation (PowerShell)

**Recommended:** Use the automated PowerShell script:

```powershell
# Navigate to script directory
cd "C:\path\to\farah-jml-script"

# Install macro
.\Install-Macro.ps1

# To uninstall later
.\Install-Macro.ps1 -Uninstall
```

**Note:** Run PowerShell as Administrator. If you get execution policy errors, run:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Manual Installation

#### Step 1: Open VBA Editor
1. Open Microsoft Word
2. Press `Alt + F11` to open the Visual Basic for Applications (VBA) Editor
3. If you see a security warning, click "Enable Macros" if prompted

#### Step 2: Access Normal Template
1. In the VBA Editor, look at the **Project Explorer** pane (usually on the left)
2. Find and expand **Normal** project
3. Expand **Microsoft Word Objects**
4. Double-click on **ThisDocument**

#### Step 3: Add Macro Code
1. Open `learner-initials-macro.bas` in a text editor (Notepad)
2. Copy the entire contents (excluding the first line if it starts with "Attribute")
3. In the VBA Editor's code window (right pane), paste the code
4. Ensure the code window shows `ThisDocument` in the dropdown at the top

#### Step 4: Save & Close
1. The code should contain a `Private Sub Document_Open()` function
2. Close the VBA Editor (`Alt + Q` or File → Close)
3. Save any changes if prompted

### Step 5: Test the Macro

**Test with an existing document:**

1. Open an existing Word document (or create a new one)
2. Type: "Jared Learner" or "JL" somewhere in the document (exact casing)
3. Save and close the document
4. Reopen the document
5. You should see a popup asking "Include Mr. Learner's initial?"
6. Click "Yes" to test the replacement

**Test with a template:**

1. Open a template file (.dotm or .dotx) directly
2. If it contains "Jared Learner" or "JL", the macro will trigger
3. The macro works the same way for templates as regular documents

**Important**: The macro uses whole-word matching. Phrases can appear at the start/end of a line or be followed by punctuation (e.g., "Jared Learner," or "(JL)") and will still be detected.

## Method 2: Custom Template Installation (Safer Alternative)

This method creates a custom template that only affects documents created from that template.

### Step 1: Create a New Template

1. Open Microsoft Word
2. Press `Alt + F11` to open VBA Editor
3. In Project Explorer, right-click on your document project
4. Select **Insert** → **Module** (or add code to ThisDocument)

### Step 2: Add the Macro Code

1. Copy the contents of `learner-initials-macro.bas`
2. Paste into the code window
3. Ensure `Document_Open` is a private subroutine

### Step 3: Save as Template

1. Close VBA Editor
2. In Word, go to **File** → **Save As**
3. Choose **Word Macro-Enabled Template (\*.dotm)**
4. Save as `LearnerInitialsTemplate.dotm`
5. Save location: `%APPDATA%\Microsoft\Templates` (Word's default template folder)

### Step 4: Use the Template

1. When creating new documents, use **File** → **New** → **Personal** → Select your template
2. Documents created from this template will have the macro functionality

## Macro Security Settings

For the macro to run, you need to enable macros in Word:

### Step 1: Open Trust Center

1. In Word, go to **File** → **Options**
2. Select **Trust Center**
3. Click **Trust Center Settings...**

### Step 2: Configure Macro Settings

1. Select **Macro Settings** from the left sidebar
2. Choose one of these options:

   - **"Disable all macros with notification"** (recommended) - Shows a security warning when macros run
   - **"Enable all macros"** (less secure) - Runs all macros automatically

3. Click **OK** to save

### Step 3: Trust Your Template (If Needed)

If using Method 2 (Custom Template):

1. In Trust Center → **Trusted Locations**
2. Ensure `%APPDATA%\Microsoft\Templates` is listed as a trusted location
3. If not, click **Add new location** and add it

## Verification Checklist

After installation, verify:

- [ ] Macro code is visible in VBA Editor
- [ ] Macro security is configured appropriately
- [ ] Test document with "Jared Learner" triggers the popup
- [ ] Test document with "JL" triggers the popup
- [ ] Clicking "Yes" performs the replacements correctly
- [ ] Clicking "No" closes without changes
- [ ] "Done!" confirmation appears after replacement

## Troubleshooting

### Problem: Macro doesn't run when document opens

**Solutions:**

- Verify macros are enabled in Trust Center
- Check that code is in `ThisDocument` module (not a separate module)
- Ensure the subroutine is named exactly `Document_Open()` (case-sensitive)
- Make sure the document isn't read-only
- Close and reopen Word

### Problem: "Macros are disabled" error

**Solutions:**

- Go to File → Options → Trust Center → Trust Center Settings
- Adjust Macro Settings (see Macro Security Settings above)
- Restart Word

### Problem: Macro runs but doesn't find text

**Solutions:**

- Verify the text is exactly "Jared Learner" or "JL" (uppercase J and L)
- Ensure the phrases are not misspelled (e.g., "Jared Learners" will not match)
- Whole-word matching allows punctuation: "Jared Learner," or "(JL)" should still work
- Check for hidden characters or unusual formatting that might interfere
- If Track Changes is enabled, ensure the text isn't marked for deletion before testing

### Problem: Normal.dotm won't save

**Solutions:**

- Close all Word documents completely
- Check if Word is running in Task Manager and end process
- Run Word as Administrator
- Ensure you have write permissions to the Templates folder

### Problem: Want to remove the macro

**Solutions:**

1. Open VBA Editor (Alt + F11)
2. Navigate to Normal → Microsoft Word Objects → ThisDocument
3. Delete the `Document_Open` subroutine code
4. Save and close
5. Or restore your Normal.dotm.backup file

## Uninstallation

To remove the macro:

1. Open VBA Editor (`Alt + F11`)
2. Navigate to **Normal** → **Microsoft Word Objects** → **ThisDocument**
3. Delete the `Document_Open` subroutine code
4. Close VBA Editor
5. Or restore your Normal.dotm backup

## Additional Notes

- The macro searches for exact matches: "Jared Learner" and "JL" (case-sensitive)
- Whole-word matching is used so occurrences at the start/end of a paragraph or next to punctuation are still caught
- Lowercase variations ("jared learner" or "jl") will NOT match
- Whole-word matching prevents false positives (e.g., won't match "AJL" or "MJared Learner")
- The macro only runs when documents are opened, not on every edit
- Multiple replacements in one document are handled automatically
- **Works with all document types**: new documents, existing documents (.docx, .doc), and templates (.dotm, .dotx)
- When you open an existing document or template, the macro automatically checks it
- Templates opened for editing will trigger the macro; documents created FROM templates will also trigger it when opened

## Support

If you encounter issues not covered here:

1. Verify all steps were followed correctly
2. Check Word version compatibility (works with Word 2010 and later)
3. Ensure VBA is installed (usually included with Word)

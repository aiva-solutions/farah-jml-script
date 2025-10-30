# Word VBA Macro: Auto-Add Jared M. Learner Initials

## Overview

This VBA macro automatically triggers when any Word document opens and searches for the phrases "Jared Learner" and "JL" (case-sensitive). If found, it prompts the user to add Mr. Learner's middle initial, then performs the replacements if confirmed.

## Features

- **Auto-triggers** on document open (no manual activation needed)
- **Works with all document types**:
  - New documents
  - Existing documents (.docx, .doc)
  - Template files (.dotm, .dotx)
- **Case-sensitive** search for standalone "Jared Learner" and "JL" (matches even at start/end of paragraphs)
- **User confirmation** via popup dialog before making changes
- **Smart replacements**:
  - "Jared Learner" → "Jared M. Learner"
  - "JL" → "JML"
- **Confirmation message** after successful replacement

## Requirements

- Microsoft Word (Windows version)
- Macros enabled in Word security settings
- Administrator access (for Normal.dotm installation method)
- **No additional PowerShell packages needed** - script uses built-in Windows PowerShell features

## Installation

**📥 New Users:** Start with [INSTALL-INSTRUCTIONS.md](INSTALL-INSTRUCTIONS.md) - Complete step-by-step guide from download to installation  
**⚡ Quick Reference:** See [QUICK-INSTALL.md](QUICK-INSTALL.md) for abbreviated commands  
**🔧 Manual Install:** See [INSTALLATION.md](INSTALLATION.md) for manual VBA Editor method

Two installation methods are available:
1. **Automated PowerShell Script** - One command installation (recommended)
2. **Manual VBA Editor Method** - Step-by-step guide for manual installation

## Usage

Once installed, the macro will automatically work with **any Word document** that opens:
- **New documents** created from scratch
- **Existing documents** (.docx, .doc files) opened from disk
- **Template files** (.dotm, .dotx) opened for editing

When a document opens, the macro will:
1. Check the document content when it opens
2. Search for "Jared Learner" or "JL" (case-sensitive, treated as standalone words even at boundaries)
3. If found, show a popup: "Include Mr. Learner's initial?"
4. If user clicks "Yes":
   - Replace all "Jared Learner" with "Jared M. Learner"
   - Replace all "JL" with "JML"
   - Show "Done!" confirmation
5. If user clicks "No": Close without making changes

**Note**: The macro only runs when documents are opened, not when they are created from templates. For templates themselves, the macro runs when you open the template file directly.

**Search Accuracy**: The macro uses Word's whole-word matching to avoid false positives inside other words, while still catching phrases without surrounding spaces.

## Important Notes

- The macro is **case-sensitive**: "jared learner" or "jl" will NOT be matched
- Only exact-case standalone instances of "Jared Learner" and "JL" are matched
- Whole-word matching prevents false positives inside other words (e.g., won't match "AJL" or "MJared Learner")
- The macro runs automatically - no button clicking required
- Make sure macros are enabled in Word's Trust Center settings

## Safety Considerations

- **Backup Normal.dotm** before installation (recommended)
- Test on a sample document first
- The macro only runs when documents are opened (not on every keystroke)
- User confirmation prevents accidental changes

## Troubleshooting

If the macro doesn't work:
1. Ensure macros are enabled in Word
2. Check that the macro is installed in the correct location (Normal.dotm or template)
3. Verify the document is not read-only
4. See INSTALLATION.md for detailed troubleshooting steps

## Support

For installation issues or questions, refer to the INSTALLATION.md guide.


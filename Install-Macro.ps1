# PowerShell Script: Install Learner Initials Macro
# This script automatically installs the VBA macro to Normal.dotm
# Run with: .\Install-Macro.ps1

param(
    [switch] $Uninstall
)

$ErrorActionPreference = "Stop"

# Get macro file path
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$macroFile = Join-Path $scriptDir "learner-initials-macro.bas"

if (-not $Uninstall) {
    # Installation
    Write-Host "Installing Learner Initials Macro..." -ForegroundColor Green
    Write-Host ""
    
    if (-not (Test-Path $macroFile)) {
        Write-Host "ERROR: Cannot find $macroFile" -ForegroundColor Red
        Write-Host "Please ensure learner-initials-macro.bas is in the same directory as this script." -ForegroundColor Yellow
        exit 1
    }
    
    # Read macro code (skip Attribute line)
    $macroCode = Get-Content $macroFile | Where-Object { $_ -notmatch "^Attribute" }
    
    try {
        # Close any open Word instances
        Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        
        # Create Word application object
        Write-Host "Opening Microsoft Word..." -ForegroundColor Cyan
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        
        # Get Normal template
        Write-Host "Accessing Normal template..." -ForegroundColor Cyan
        $normalTemplate = $word.NormalTemplate
        
        # Open VBA Editor and add code
        Write-Host "Installing macro code..." -ForegroundColor Cyan
        $normalTemplate.OpenAsDocument()
        $doc = $word.ActiveDocument
        
        # Get VBA project
        $vbaProject = $normalTemplate.VBProject
        $vbComponent = $vbaProject.VBComponents("ThisDocument")
        
        # Check if macro already exists
        $existingCode = $vbComponent.CodeModule.Lines(1, $vbComponent.CodeModule.CountOfLines)
        if ($existingCode -match "Document_Open.*Jared Learner") {
            Write-Host "Macro already exists. Updating..." -ForegroundColor Yellow
            # Find and remove existing Document_Open function
            $module = $vbComponent.CodeModule
            $startLine = 1
            $endLine = $module.CountOfLines
            
            for ($i = 1; $i -le $endLine; $i++) {
                $line = $module.Lines($i, 1)
                if ($line -match "Private Sub Document_Open\(\)") {
                    $startLine = $i
                    # Find End Sub
                    for ($j = $i; $j -le $endLine; $j++) {
                        if ($module.Lines($j, 1) -match "^\s*End Sub\s*$") {
                            $endLine = $j
                            break
                        }
                    }
                    # Delete the function
                    $module.DeleteLines($startLine, $endLine - $startLine + 1)
                    break
                }
            }
        }
        
        # Add new macro code
        $vbComponent.CodeModule.AddFromString($macroCode)
        
        # Save and close
        $normalTemplate.Save()
        $word.Quit()
        
        Write-Host ""
        Write-Host "✓ Macro installed successfully!" -ForegroundColor Green
        Write-Host ""
        Write-Host "The macro will now run automatically when any Word document opens." -ForegroundColor Cyan
        Write-Host "To test: Create a Word document with 'Jared Learner' or 'JL' and reopen it." -ForegroundColor Cyan
        
        # Release COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
    } catch {
        Write-Host ""
        Write-Host "ERROR: Installation failed" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-Host ""
        Write-Host "Try manual installation: See INSTALLATION.md" -ForegroundColor Yellow
        
        if ($word) {
            try { $word.Quit() } catch {}
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        }
        exit 1
    }
    
} else {
    # Uninstallation
    Write-Host "Uninstalling Learner Initials Macro..." -ForegroundColor Yellow
    Write-Host ""
    
    try {
        # Close any open Word instances
        Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        
        # Create Word application object
        Write-Host "Opening Microsoft Word..." -ForegroundColor Cyan
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        
        # Get Normal template
        $normalTemplate = $word.NormalTemplate
        $vbaProject = $normalTemplate.VBProject
        $vbComponent = $vbaProject.VBComponents("ThisDocument")
        
        # Find and remove Document_Open function
        $module = $vbComponent.CodeModule
        $found = $false
        
        for ($i = 1; $i -le $module.CountOfLines; $i++) {
            $line = $module.Lines($i, 1)
            if ($line -match "Private Sub Document_Open\(\)") {
                $startLine = $i
                # Find End Sub
                for ($j = $i; $j -le $module.CountOfLines; $j++) {
                    if ($module.Lines($j, 1) -match "^\s*End Sub\s*$") {
                        $endLine = $j
                        $module.DeleteLines($startLine, $endLine - $startLine + 1)
                        $found = $true
                        break
                    }
                }
                break
            }
        }
        
        if ($found) {
            $normalTemplate.Save()
            Write-Host "✓ Macro uninstalled successfully!" -ForegroundColor Green
        } else {
            Write-Host "No macro found to uninstall." -ForegroundColor Yellow
        }
        
        $word.Quit()
        
        # Release COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
    } catch {
        Write-Host ""
        Write-Host "ERROR: Uninstallation failed" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        
        if ($word) {
            try { $word.Quit() } catch {}
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        }
        exit 1
    }
}


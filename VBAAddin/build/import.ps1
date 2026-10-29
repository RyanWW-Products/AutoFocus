param(
    [string]$SourceRoot = (Join-Path $PSScriptRoot "..\src"),
    [string]$OutputDir  = (Join-Path $PSScriptRoot "..\out"),
    [string]$AddinName  = "AutoFocus",
    [string[]]$Extensions = @('*.bas','*.cls','*.frm')
)

Write-Host "Building PPAM from" $SourceRoot

if (-not (Test-Path $SourceRoot)) {
    throw "SourceRoot not found: $SourceRoot"
}

if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

# Make sure PowerPoint will allow programmatic access to VBProject (user setting)
Write-Host "Reminder: In PowerPoint, enable 'Trust access to the VBA project object model'."

# Start PowerPoint via COM
$pp = $null
$pres = $null
try {
    $pp = New-Object -ComObject PowerPoint.Application
    # Optional: show window while building (MsoTriState: -1 = True, 0 = False)
    $pp.Visible = -1

    $pres = $pp.Presentations.Add()

    # Verify VBA project access (Trust Center setting)
    try {
        $null = $pres.VBProject.VBComponents.Count
    }
    catch {
        throw "PowerPoint is blocking access to the VBA project. Enable 'Trust access to the VBA project object model' in Trust Center, restart PowerPoint, then rerun."
    }

    function Import-Path([string]$path) {
        if (Test-Path $path) {
            try {
                # Imports .bas/.cls/.frm (for .frm, keep matching .frx beside it)
                $null = $pres.VBProject.VBComponents.Import($path)
                Write-Host "Imported" (Split-Path $path -Leaf)
            }
            catch {
                Write-Warning "Failed to import ${path}: $_"
            }
        }
    }

    # Import all modules, classes, and forms (recursive)
    Get-ChildItem -Path $SourceRoot -Recurse -Include $Extensions -File | Sort-Object FullName | ForEach-Object {
        Import-Path $_.FullName
    }

    # Save as OpenXML Add-in (.ppam)
    $output = Join-Path $OutputDir ("{0}.ppam" -f $AddinName)
    $ppSaveAsOpenXMLAddin = 30
    $pres.SaveAs($output, $ppSaveAsOpenXMLAddin)
    Write-Host "Saved add-in to" $output
}
finally {
    if ($pres -ne $null) { $pres.Close() | Out-Null }
    if ($pp -ne $null)   { $pp.Quit()  | Out-Null }

    # Release COM objects cleanly
    if ($pres -ne $null) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) }
    if ($pp -ne $null)   { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

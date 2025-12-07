# PPT2PDF.ps1
param (
    [Parameter(ValueFromRemainingArguments=$true)]
    [string[]]$Files
)

# Check input
if ($Files.Count -eq 0) {
    Write-Host "No files provided."
    exit
}

# Create PowerPoint COM Object
$pptApp = New-Object -ComObject PowerPoint.Application
# PowerPoint must be visible for the export to work reliably
$pptApp.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

# PDF format code (ppSaveAsPDF = 32)
$ppSaveAsPDF = 32

foreach ($file in $Files) {
    # absolute path
    $filePath = (Resolve-Path $file).Path
    
    # Convert file name through reg
    $pdfPath = $filePath -replace '\.pptx?$', '.pdf'
    
    Write-Host "Processing: $filePath ..."
    
    # Check if the file exists
    if (Test-Path -Path $pdfPath) {
        Remove-Item -Path $pdfPath -Force
        Write-Host "Overwriting existing file..." -ForegroundColor Yellow
    }

    try {
        # Open Presentation (ReadOnly, Untitled, WithWindow)
        $presentation = $pptApp.Presentations.Open($filePath, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
        
        # Save as PDF
        $presentation.SaveAs($pdfPath, $ppSaveAsPDF)
        
        # Close Presentation
        $presentation.Close()
        Write-Host "Saved: $pdfPath" -ForegroundColor Green
    }
    catch {
        Write-Host "Error converting: $file" -ForegroundColor Red
        Write-Host $_.Exception.Message
    }
}

# Quit PowerPoint
$pptApp.Quit()

# Cleanup COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($pptApp) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Host "Done. Closing in 3 seconds..."
Start-Sleep -Seconds 3
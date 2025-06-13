# PowerShell Script to Open Excel and Run a Macro

$excelFilePath = "C:\Path\To\YourWorkbook.xlsm"
$macroName = "Module1.MyMacro"  # format: ModuleName.MacroName

# Start Excel COM object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# Open the workbook
$workbook = $excel.Workbooks.Open($excelFilePath)

try {
    # Run the macro
    $excel.Run($macroName)
    
    # Save and close
    $workbook.Save()
    $workbook.Close($false)
} catch {
    Write-Host "Error running macro: $_"
} finally {
    # Quit Excel and release COM object
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Remove-Variable excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

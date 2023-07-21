param (
        [Parameter(Mandatory = $true)]
        [string]$InputPath,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputPath
)

    # Create an Excel COM Object
    $excel = New-Object -ComObject Excel.Application

    # Don't display Excel on the screen
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # Open the xls file
    $workbook = $excel.Workbooks.Open($InputPath)

    # Save it in the xlsx format
    $xlsxFormat = 51 # This is the code for xlsx format
    $workbook.SaveAs($OutputPath, $xlsxFormat)

    # Clean up
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
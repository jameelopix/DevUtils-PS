function Convert-CsvToXlsx {

    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$InputFile,

        [Parameter(Position = 1)]
        [string]$OutputFile
    )

    # Resolve input file
    $InputFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath(
        $InputFile
    )

    if (!(Test-Path $InputFile)) {
        throw "Input file not found: $InputFile"
    }

    if ([System.IO.Path]::GetExtension($InputFile) -ne ".csv") {
        throw "Input file must be a CSV file."
    }

    # Generate output filename if not supplied
    if ([string]::IsNullOrWhiteSpace($OutputFile)) {

        $Directory = Split-Path $InputFile
        $FileName = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)

        $OutputFile = Join-Path $Directory "$FileName.xlsx"
    }
    else {
        $OutputFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath(
            $OutputFile
        )
    }

    # Check dependency
    if (!(Get-Module -ListAvailable -Name ImportExcel)) {
        throw @"
ImportExcel PowerShell module is required.

Install it using:

Install-Module ImportExcel -Scope CurrentUser

Administrator access is not required.
"@
    }

    Write-Host ""
    Write-Host "CSV -> XLSX"
    Write-Host "-----------------------------"
    Write-Host "Input : $InputFile"
    Write-Host "Output: $OutputFile"
    Write-Host ""

    try {

        $Data = Import-Csv -Path $InputFile

        if ($null -eq $Data -or $Data.Count -eq 0) {
            throw "CSV file contains no data."
        }

        $Data | Export-Excel `
            -Path $OutputFile `
            -WorksheetName "Data" `
            -AutoSize `
            -AutoFilter `
            -FreezeTopRow `
            -BoldTopRow

        Write-Host "Conversion completed successfully."
        Write-Host ""
        Write-Host $OutputFile
    }
    catch {
        throw "CSV to XLSX conversion failed: $($_.Exception.Message)"
    }
}

Export-ModuleMember -Function Convert-CsvToXlsx
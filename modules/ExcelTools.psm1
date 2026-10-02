function Test-ImportExcelDependency {
    if (!(Get-Module -ListAvailable -Name ImportExcel)) {
        throw @"
ImportExcel PowerShell module is required for Excel utilities.

Install it using:

Install-Module ImportExcel -Scope CurrentUser
"@
    }
}

function Resolve-InputFilePath {
    param([Parameter(Mandatory = $true)][string]$InputFile)

    try {
        $ResolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($InputFile)
    }
    catch {
        throw "Input file not found: $InputFile"
    }

    if (!(Test-Path -LiteralPath $ResolvedPath -PathType Leaf)) {
        throw "Input file not found: $ResolvedPath"
    }

    return $ResolvedPath
}

function New-DefaultOutputFilePath {
    param([Parameter(Mandatory = $true)][string]$InputFile)

    $Directory = Split-Path -Path $InputFile
    $FileName = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)
    return (Join-Path $Directory "$FileName-date-format.xlsx")
}

function Convert-ExcelColumnNameToNumber {
    param([Parameter(Mandatory = $true)][string]$ColumnName)

    $ColumnName = $ColumnName.ToUpperInvariant()
    $ColumnNumber = 0

    foreach ($Character in $ColumnName.ToCharArray()) {
        if ($Character -lt 'A' -or $Character -gt 'Z') {
            throw "Invalid Excel column letter: $ColumnName"
        }

        $ColumnNumber = ($ColumnNumber * 26) + ([int][char]$Character - [int][char]'A' + 1)
    }

    return $ColumnNumber
}

function Resolve-ExcelColumnNumber {
    param(
        [Parameter(Mandatory = $true)]
        $Worksheet,

        [Parameter(Mandatory = $true)]
        [string]$Column
    )

    if ($Column -match "^[A-Za-z]+$") {
        $HeaderColumn = $null

        if ($Worksheet.Dimension) {
            for ($Index = 1; $Index -le $Worksheet.Dimension.End.Column; $Index++) {
                if ([string]$Worksheet.Cells[1, $Index].Value -eq $Column) {
                    $HeaderColumn = $Index
                    break
                }
            }
        }

        if ($null -ne $HeaderColumn) {
            return $HeaderColumn
        }

        return (Convert-ExcelColumnNameToNumber $Column)
    }

    if (!$Worksheet.Dimension) {
        throw "Worksheet is empty."
    }

    for ($Index = 1; $Index -le $Worksheet.Dimension.End.Column; $Index++) {
        if ([string]$Worksheet.Cells[1, $Index].Value -eq $Column) {
            return $Index
        }
    }

    throw "Column not found: $Column"
}

function ConvertTo-DateTimeValue {
    param(
        [AllowNull()]
        $Value,

        [Parameter(Mandatory = $true)]
        [string]$SourceFormat
    )

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [datetime]) {
        return $Value
    }

    if ($Value -is [double] -or $Value -is [int] -or $Value -is [decimal]) {
        try {
            return [datetime]::FromOADate([double]$Value)
        }
        catch {
            throw "Value '$Value' is not a valid Excel date serial."
        }
    }

    $Text = ([string]$Value).Trim()

    if ($Text -eq "") {
        return $null
    }

    try {
        return [datetime]::ParseExact(
            $Text,
            $SourceFormat,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::None
        )
    }
    catch {
        try {
            return [datetime]::Parse(
                $Text,
                [System.Globalization.CultureInfo]::InvariantCulture
            )
        }
        catch {
            throw "Value '$Text' does not match source format '$SourceFormat'."
        }
    }
}

function Convert-ExcelColumnDateFormat {
    param(
        [Parameter(ValueFromRemainingArguments = $true)]
        [string[]]$Arguments
    )

    $ParsedArguments = ConvertFrom-DateFormatArguments $Arguments
    $InputFile = $ParsedArguments.InputFile
    $Column = $ParsedArguments.Column
    $SourceFormat = $ParsedArguments.SourceFormat
    $TargetFormat = $ParsedArguments.TargetFormat
    $OutputFile = $ParsedArguments.OutputFile
    $WorksheetName = $ParsedArguments.WorksheetName

    if ([string]::IsNullOrWhiteSpace($InputFile)) {
        throw "Missing input file. Use: devutils excel date-format --input-file <input.xlsx> --column <column> --from-format <format> --to-format <format>"
    }

    if ([string]::IsNullOrWhiteSpace($Column)) {
        throw "Missing column. Use a header name or Excel column letter, such as Date or C."
    }

    if ([string]::IsNullOrWhiteSpace($SourceFormat)) {
        throw "Missing source date format, such as dd/MM/yyyy."
    }

    if ([string]::IsNullOrWhiteSpace($TargetFormat)) {
        throw "Missing target date format, such as yyyy-MM-dd."
    }

    Test-ImportExcelDependency

    $InputFile = Resolve-InputFilePath $InputFile

    if ([System.IO.Path]::GetExtension($InputFile).ToLowerInvariant() -ne ".xlsx") {
        throw "Input file must be an .xlsx file."
    }

    if ([string]::IsNullOrWhiteSpace($OutputFile)) {
        $OutputFile = New-DefaultOutputFilePath $InputFile
    }
    else {
        $OutputFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFile)
    }

    Copy-Item -LiteralPath $InputFile -Destination $OutputFile -Force

    $Package = $null

    try {
        $Package = Open-ExcelPackage -Path $OutputFile

        if ([string]::IsNullOrWhiteSpace($WorksheetName)) {
            $Worksheet = $Package.Workbook.Worksheets[1]
        }
        else {
            $Worksheet = $Package.Workbook.Worksheets[$WorksheetName]
        }

        if ($null -eq $Worksheet) {
            throw "Worksheet not found: $WorksheetName"
        }

        if (!$Worksheet.Dimension -or $Worksheet.Dimension.End.Row -lt 2) {
            throw "Worksheet contains no data rows."
        }

        $ColumnNumber = Resolve-ExcelColumnNumber -Worksheet $Worksheet -Column $Column
        $UpdatedCount = 0
        $SkippedBlankCount = 0

        for ($Row = 2; $Row -le $Worksheet.Dimension.End.Row; $Row++) {
            $Cell = $Worksheet.Cells[$Row, $ColumnNumber]
            $DateValue = ConvertTo-DateTimeValue -Value $Cell.Value -SourceFormat $SourceFormat

            if ($null -eq $DateValue) {
                $SkippedBlankCount++
                continue
            }

            $Cell.Value = $DateValue
            $Cell.Style.Numberformat.Format = $TargetFormat
            $UpdatedCount++
        }

        if ($UpdatedCount -eq 0) {
            throw "No date values were updated in column '$Column'."
        }

        Close-ExcelPackage $Package
        $Package = $null

        Write-Host ""
        Write-Host "Excel Date Format"
        Write-Host "-----------------------------"
        Write-Host "Input       : $InputFile"
        Write-Host "Output      : $OutputFile"
        Write-Host "Worksheet   : $($Worksheet.Name)"
        Write-Host "Column      : $Column"
        Write-Host "From format : $SourceFormat"
        Write-Host "To format   : $TargetFormat"
        Write-Host "Updated     : $UpdatedCount"
        Write-Host "Blank cells : $SkippedBlankCount"
        Write-Host ""
        Write-Host $OutputFile
    }
    catch {
        if ($null -ne $Package) {
            Close-ExcelPackage $Package -NoSave
        }

        throw "Excel date format conversion failed: $($_.Exception.Message)"
    }
}

function ConvertFrom-DateFormatArguments {
    param([string[]]$Arguments)

    $Result = [ordered]@{
        InputFile = $null
        Column = $null
        SourceFormat = $null
        TargetFormat = $null
        OutputFile = $null
        WorksheetName = $null
    }

    if ($null -eq $Arguments -or $Arguments.Count -eq 0) {
        return [pscustomobject]$Result
    }

    if (!$Arguments[0].StartsWith("--")) {
        throw "Positional arguments are not supported. Use named options, such as: devutils excel date-format --input-file report.xlsx --column Date --from-format dd/MM/yyyy --to-format yyyy-MM-dd"
    }

    $OptionMap = @{
        "--input-file" = "InputFile"
        "--column" = "Column"
        "--from-format" = "SourceFormat"
        "--source-format" = "SourceFormat"
        "--to-format" = "TargetFormat"
        "--target-format" = "TargetFormat"
        "--output-file" = "OutputFile"
        "--sheet-name" = "WorksheetName"
        "--worksheet" = "WorksheetName"
    }

    for ($Index = 0; $Index -lt $Arguments.Count; $Index++) {
        $Option = $Arguments[$Index].ToLowerInvariant()

        if (!$OptionMap.ContainsKey($Option)) {
            throw "Unknown option: $($Arguments[$Index])"
        }

        if ($Index + 1 -ge $Arguments.Count -or $Arguments[$Index + 1].StartsWith("--")) {
            throw "Missing value for option: $($Arguments[$Index])"
        }

        $Result[$OptionMap[$Option]] = $Arguments[$Index + 1]
        $Index++
    }

    return [pscustomobject]$Result
}

Export-ModuleMember -Function Convert-ExcelColumnDateFormat

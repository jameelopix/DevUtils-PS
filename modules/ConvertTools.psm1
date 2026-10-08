function Test-ImportExcelDependency {
    if (!(Get-Module -ListAvailable -Name ImportExcel)) {
        throw @"
ImportExcel PowerShell module is required for XLSX conversions.

Install it for the current user with:

Install-Module ImportExcel -Scope CurrentUser

Administrator access is not required.
"@
    }
}

function Test-YamlDependency {
    if (!(Get-Module -ListAvailable -Name powershell-yaml)) {
        throw @"
powershell-yaml PowerShell module is required for YAML conversions.

Install it for the current user with:

Install-Module powershell-yaml -Scope CurrentUser

Administrator access is not required.
"@
    }

    Import-Module powershell-yaml -ErrorAction Stop
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

function Get-NormalizedFormat {
    param([Parameter(Mandatory = $true)][string]$Format)

    return $Format.Trim().TrimStart(".").ToLowerInvariant()
}

function Get-TargetExtension {
    param([Parameter(Mandatory = $true)][string]$TargetFormat)

    switch (Get-NormalizedFormat $TargetFormat) {
        "xlsx" { return "xlsx" }
        "csv" { return "csv" }
        "json" { return "json" }
        "yaml" { return "yaml" }
        "yml" { return "yaml" }
        "xml" { return "xml" }
        default { throw "Unsupported target format: $TargetFormat" }
    }
}

function Resolve-OutputFilePath {
    param(
        [Parameter(Mandatory = $true)][string]$InputFile,
        [Parameter(Mandatory = $true)][string]$TargetFormat,
        [string]$OutputFile,
        [bool]$Force
    )

    $Extension = Get-TargetExtension $TargetFormat

    if ([string]::IsNullOrWhiteSpace($OutputFile)) {
        $Directory = Split-Path -Path $InputFile
        $FileName = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)
        $OutputFile = Join-Path $Directory "$FileName.$Extension"
    }
    else {
        $OutputFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFile)
    }

    $OutputDirectory = Split-Path -Path $OutputFile

    if (![string]::IsNullOrWhiteSpace($OutputDirectory) -and !(Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
        throw "Output directory does not exist: $OutputDirectory"
    }

    if ((Test-Path -LiteralPath $OutputFile) -and !$Force) {
        throw "Output file already exists: $OutputFile. Use --force to overwrite it."
    }

    return $OutputFile
}

function Read-OptionValue {
    param(
        [Parameter(Mandatory = $true)][string[]]$Arguments,
        [Parameter(Mandatory = $true)][ref]$Index,
        [Parameter(Mandatory = $true)][string]$Option
    )

    if ($Index.Value + 1 -ge $Arguments.Count -or $Arguments[$Index.Value + 1].StartsWith("--")) {
        throw "Missing value for option: $Option"
    }

    $Index.Value++
    return $Arguments[$Index.Value]
}

function ConvertFrom-ConvertArguments {
    param([string[]]$Arguments)

    $Result = [ordered]@{
        InputFile = $null
        SourceFormat = $null
        TargetFormat = $null
        OutputFile = $null
        Force = $false
        Help = $false
    }

    if ($null -eq $Arguments -or $Arguments.Count -eq 0) {
        return [pscustomobject]$Result
    }

    for ($Index = 0; $Index -lt $Arguments.Count; $Index++) {
        $Argument = $Arguments[$Index]

        if ($Argument -eq "--help" -or $Argument -eq "-h") {
            $Result.Help = $true
            continue
        }

        if ($Argument -eq "--to") {
            $Result.TargetFormat = Read-OptionValue -Arguments $Arguments -Index ([ref]$Index) -Option $Argument
            continue
        }

        if ($Argument -eq "--from") {
            $Result.SourceFormat = Read-OptionValue -Arguments $Arguments -Index ([ref]$Index) -Option $Argument
            continue
        }

        if ($Argument -eq "--output") {
            $Result.OutputFile = Read-OptionValue -Arguments $Arguments -Index ([ref]$Index) -Option $Argument
            continue
        }

        if ($Argument -eq "--force") {
            $Result.Force = $true
            continue
        }

        if ($Argument.StartsWith("--")) {
            throw "Unknown option: $Argument"
        }

        if ($null -ne $Result.InputFile) {
            throw "Unexpected argument: $Argument"
        }

        $Result.InputFile = $Argument
    }

    return [pscustomobject]$Result
}

function Resolve-SourceFormat {
    param(
        [Parameter(Mandatory = $true)][string]$InputFile,
        [string]$ExplicitSourceFormat
    )

    if (![string]::IsNullOrWhiteSpace($ExplicitSourceFormat)) {
        return (Get-NormalizedFormat $ExplicitSourceFormat)
    }

    $Extension = [System.IO.Path]::GetExtension($InputFile)

    if ([string]::IsNullOrWhiteSpace($Extension)) {
        throw "Unable to infer source format because the input file has no extension. Use --from <format>."
    }

    $Format = Get-NormalizedFormat $Extension

    if ($Format -eq "txt") {
        throw "Ambiguous source extension '.txt'. Use --from <format>, for example --from csv."
    }

    return $Format
}

function Get-ConversionMap {
    return @{
        "csv->xlsx" = { param($Source, $Destination) Convert-CsvToXlsx -InputFile $Source -OutputFile $Destination }
        "xlsx->csv" = { param($Source, $Destination) Convert-XlsxToCsv -InputFile $Source -OutputFile $Destination }
        "csv->json" = { param($Source, $Destination) Convert-CsvToJson -InputFile $Source -OutputFile $Destination }
        "json->csv" = { param($Source, $Destination) Convert-JsonToCsv -InputFile $Source -OutputFile $Destination }
        "json->yaml" = { param($Source, $Destination) Convert-JsonToYaml -InputFile $Source -OutputFile $Destination }
        "yaml->json" = { param($Source, $Destination) Convert-YamlToJson -InputFile $Source -OutputFile $Destination }
        "yml->json" = { param($Source, $Destination) Convert-YamlToJson -InputFile $Source -OutputFile $Destination }
        "xml->json" = { param($Source, $Destination) Convert-XmlToJson -InputFile $Source -OutputFile $Destination }
        "json->xml" = { param($Source, $Destination) Convert-JsonToXml -InputFile $Source -OutputFile $Destination }
    }
}

function Test-SupportedFormat {
    param(
        [Parameter(Mandatory = $true)][string]$SourceFormat,
        [Parameter(Mandatory = $true)][string]$TargetFormat
    )

    $SupportedSources = @("csv", "xlsx", "json", "yaml", "yml", "xml")
    $SupportedTargets = @("xlsx", "csv", "json", "yaml", "yml", "xml")

    if ($SourceFormat -notin $SupportedSources) {
        throw "Unsupported source format: $SourceFormat. Supported sources: $($SupportedSources -join ', ')."
    }

    if ($TargetFormat -notin $SupportedTargets) {
        throw "Unsupported target format: $TargetFormat. Supported targets: $($SupportedTargets -join ', ')."
    }
}

function Read-CsvRows {
    param([Parameter(Mandatory = $true)][string]$InputFile)

    $Rows = Import-Csv -LiteralPath $InputFile

    if ($null -eq $Rows -or $Rows.Count -eq 0) {
        throw "CSV file contains no data rows."
    }

    return @($Rows)
}

function Read-JsonDocument {
    param([Parameter(Mandatory = $true)][string]$InputFile)

    $Content = Get-Content -LiteralPath $InputFile -Raw -Encoding UTF8

    if ([string]::IsNullOrWhiteSpace($Content)) {
        throw "JSON file contains no data."
    }

    try {
        $Parsed = $Content | ConvertFrom-Json

        if ($Content.TrimStart().StartsWith("[") -and $Parsed -isnot [System.Collections.IList]) {
            return ,@($Parsed)
        }

        return $Parsed
    }
    catch {
        throw "Invalid JSON content: $($_.Exception.Message)"
    }
}

function Read-YamlDocument {
    param([Parameter(Mandatory = $true)][string]$InputFile)

    Test-YamlDependency
    $Content = Get-Content -LiteralPath $InputFile -Raw -Encoding UTF8

    if ([string]::IsNullOrWhiteSpace($Content)) {
        throw "YAML file contains no data."
    }

    try {
        return (ConvertFrom-Yaml -Yaml $Content)
    }
    catch {
        throw "Invalid YAML content: $($_.Exception.Message)"
    }
}

function Write-JsonDocument {
    param(
        [Parameter(Mandatory = $true)]$Value,
        [Parameter(Mandatory = $true)][string]$OutputFile
    )

    $Value | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $OutputFile -Encoding UTF8
}

function Write-YamlDocument {
    param(
        [Parameter(Mandatory = $true)]$Value,
        [Parameter(Mandatory = $true)][string]$OutputFile
    )

    Test-YamlDependency
    $Value | ConvertTo-Yaml | Set-Content -LiteralPath $OutputFile -Encoding UTF8
}

function ConvertTo-PlainObject {
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return $null }

    if ($Value -is [System.Collections.IDictionary]) {
        $Result = [ordered]@{}
        foreach ($Key in $Value.Keys) {
            $Result[$Key] = ConvertTo-PlainObject $Value[$Key]
        }
        return $Result
    }

    if ($Value -is [pscustomobject]) {
        $Result = [ordered]@{}
        foreach ($Property in $Value.PSObject.Properties) {
            $Result[$Property.Name] = ConvertTo-PlainObject $Property.Value
        }
        return $Result
    }

    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        $Items = @()
        foreach ($Item in $Value) {
            $Items += ConvertTo-PlainObject $Item
        }
        return $Items
    }

    return $Value
}

function Test-FlatObjectArray {
    param([Parameter(Mandatory = $true)]$Value)

    if ($Value -is [pscustomobject] -or $Value -is [System.Collections.IDictionary]) {
        $Rows = @($Value)
    }
    elseif ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        $Rows = @($Value)
    }
    else {
        throw "JSON to CSV supports arrays of flat objects only."
    }

    if ($Rows.Count -eq 0) {
        throw "JSON array contains no rows."
    }

    foreach ($Row in $Rows) {
        if ($Row -isnot [pscustomobject] -and $Row -isnot [System.Collections.IDictionary]) {
            throw "JSON to CSV supports arrays of flat objects only."
        }

        $PlainRow = ConvertTo-PlainObject $Row

        foreach ($Key in $PlainRow.Keys) {
            $CellValue = $PlainRow[$Key]
            if ($CellValue -is [System.Collections.IDictionary] -or ($CellValue -is [System.Collections.IEnumerable] -and $CellValue -isnot [string])) {
                throw "JSON to CSV does not support nested property '$Key'. Use a flat array of objects."
            }
        }
    }

    return $Rows
}

function Convert-XmlNodeToObject {
    param([Parameter(Mandatory = $true)][System.Xml.XmlNode]$Node)

    $ElementChildren = @($Node.ChildNodes | Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element })

    if ($Node.Attributes.Count -eq 0 -and $ElementChildren.Count -eq 0) {
        return $Node.InnerText
    }

    $Result = [ordered]@{}

    foreach ($Attribute in $Node.Attributes) {
        $Result["@$(($Attribute.LocalName))"] = $Attribute.Value
    }

    foreach ($Child in $ElementChildren) {
        $ChildName = $Child.LocalName
        $ChildValue = Convert-XmlNodeToObject $Child

        if ($Result.Contains($ChildName)) {
            if ($Result[$ChildName] -isnot [System.Collections.IList]) {
                $Result[$ChildName] = @($Result[$ChildName])
            }
            $Result[$ChildName] += $ChildValue
        }
        else {
            $Result[$ChildName] = $ChildValue
        }
    }

    return $Result
}

function Add-JsonXmlElement {
    param(
        [Parameter(Mandatory = $true)][System.Xml.XmlDocument]$Document,
        [Parameter(Mandatory = $true)][System.Xml.XmlNode]$Parent,
        [Parameter(Mandatory = $true)][string]$Name,
        [AllowNull()]$Value
    )

    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string] -and $Value -isnot [System.Collections.IDictionary] -and $Value -isnot [pscustomobject]) {
        foreach ($Item in $Value) {
            Add-JsonXmlElement -Document $Document -Parent $Parent -Name $Name -Value $Item
        }
        return
    }

    $Element = $Document.CreateElement($Name)
    [void]$Parent.AppendChild($Element)
    $Value = ConvertTo-PlainObject $Value

    if ($Value -is [System.Collections.IDictionary]) {
        foreach ($Key in $Value.Keys) {
            if ($Key.StartsWith("@")) {
                $Element.SetAttribute($Key.Substring(1), [string]$Value[$Key])
            }
            else {
                Add-JsonXmlElement -Document $Document -Parent $Element -Name $Key -Value $Value[$Key]
            }
        }
    }
    elseif ($null -ne $Value) {
        $Element.InnerText = [string]$Value
    }
}

function Write-ConversionSummary {
    param(
        [Parameter(Mandatory = $true)][string]$SourceFormat,
        [Parameter(Mandatory = $true)][string]$TargetFormat,
        [Parameter(Mandatory = $true)][string]$InputFile,
        [Parameter(Mandatory = $true)][string]$OutputFile
    )

    Write-Host ""
    Write-Host "$($SourceFormat.ToUpperInvariant()) -> $($TargetFormat.ToUpperInvariant())"
    Write-Host "-----------------------------"
    Write-Host "Input : $InputFile"
    Write-Host "Output: $OutputFile"
    Write-Host ""
    Write-Host "Conversion completed successfully."
    Write-Host ""
    Write-Host $OutputFile
}

function Convert-CsvToXlsx {
    param([string]$InputFile, [string]$OutputFile)

    Test-ImportExcelDependency
    $Rows = Read-CsvRows $InputFile
    $Rows | Export-Excel -Path $OutputFile -WorksheetName "Data" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
}

function Convert-XlsxToCsv {
    param([string]$InputFile, [string]$OutputFile)

    Test-ImportExcelDependency
    $Rows = Import-Excel -Path $InputFile

    if ($null -eq $Rows -or $Rows.Count -eq 0) {
        throw "XLSX file contains no data rows."
    }

    $Rows | Export-Csv -LiteralPath $OutputFile -NoTypeInformation -Encoding UTF8
}

function Convert-CsvToJson {
    param([string]$InputFile, [string]$OutputFile)

    Write-JsonDocument -Value (Read-CsvRows $InputFile) -OutputFile $OutputFile
}

function Convert-JsonToCsv {
    param([string]$InputFile, [string]$OutputFile)

    $Rows = Test-FlatObjectArray (Read-JsonDocument $InputFile)
    $Rows | Export-Csv -LiteralPath $OutputFile -NoTypeInformation -Encoding UTF8
}

function Convert-JsonToYaml {
    param([string]$InputFile, [string]$OutputFile)

    Write-YamlDocument -Value (Read-JsonDocument $InputFile) -OutputFile $OutputFile
}

function Convert-YamlToJson {
    param([string]$InputFile, [string]$OutputFile)

    Write-JsonDocument -Value (Read-YamlDocument $InputFile) -OutputFile $OutputFile
}

function Convert-XmlToJson {
    param([string]$InputFile, [string]$OutputFile)

    try {
        [xml]$Xml = Get-Content -LiteralPath $InputFile -Raw -Encoding UTF8
    }
    catch {
        throw "Invalid XML content: $($_.Exception.Message)"
    }

    $RootName = $Xml.DocumentElement.LocalName
    $Data = [ordered]@{}
    $Data[$RootName] = Convert-XmlNodeToObject $Xml.DocumentElement
    Write-JsonDocument -Value $Data -OutputFile $OutputFile
}

function Convert-JsonToXml {
    param([string]$InputFile, [string]$OutputFile)

    $Data = ConvertTo-PlainObject (Read-JsonDocument $InputFile)
    $Xml = New-Object System.Xml.XmlDocument
    $Declaration = $Xml.CreateXmlDeclaration("1.0", "utf-8", $null)
    [void]$Xml.AppendChild($Declaration)

    if ($Data -is [System.Collections.IDictionary] -and $Data.Count -eq 1) {
        $RootName = @($Data.Keys)[0]
        Add-JsonXmlElement -Document $Xml -Parent $Xml -Name $RootName -Value $Data[$RootName]
    }
    else {
        Add-JsonXmlElement -Document $Xml -Parent $Xml -Name "root" -Value $Data
    }

    $Xml.Save($OutputFile)
}

function Invoke-FileConversion {
    param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Arguments)

    $ParsedArguments = ConvertFrom-ConvertArguments $Arguments

    if ($ParsedArguments.Help) {
        Show-FileConversionHelp
        return
    }

    if ([string]::IsNullOrWhiteSpace($ParsedArguments.InputFile)) {
        throw "Missing input file. Use: devutils convert <input> --to <target>"
    }

    if ([string]::IsNullOrWhiteSpace($ParsedArguments.TargetFormat)) {
        throw "Missing required option: --to <target>"
    }

    $InputFile = Resolve-InputFilePath $ParsedArguments.InputFile
    $SourceFormat = Resolve-SourceFormat -InputFile $InputFile -ExplicitSourceFormat $ParsedArguments.SourceFormat
    $TargetFormat = Get-NormalizedFormat $ParsedArguments.TargetFormat

    if ($TargetFormat -eq "yml") {
        $TargetFormat = "yaml"
    }

    Test-SupportedFormat -SourceFormat $SourceFormat -TargetFormat $TargetFormat

    $Route = "$SourceFormat->$TargetFormat"
    $Conversions = Get-ConversionMap

    if (!$Conversions.ContainsKey($Route)) {
        throw "Unsupported conversion: $SourceFormat -> $TargetFormat. Run 'devutils convert --help' for supported conversions."
    }

    $OutputFile = Resolve-OutputFilePath `
        -InputFile $InputFile `
        -TargetFormat $TargetFormat `
        -OutputFile $ParsedArguments.OutputFile `
        -Force $ParsedArguments.Force

    try {
        & $Conversions[$Route] $InputFile $OutputFile
        Write-ConversionSummary -SourceFormat $SourceFormat -TargetFormat $TargetFormat -InputFile $InputFile -OutputFile $OutputFile
    }
    catch {
        throw "File conversion failed: $($_.Exception.Message)"
    }
}

function Show-FileConversionHelp {
    Write-Host ""
    Write-Host "File conversion commands:"
    Write-Host "  devutils convert <input> --to <target> [--output <output>] [--force]"
    Write-Host "  devutils convert <input> --from <source> --to <target> [--output <output>] [--force]"
    Write-Host ""
    Write-Host "Examples:"
    Write-Host "  devutils convert data.csv --to xlsx"
    Write-Host "  devutils convert data.xlsx --to csv"
    Write-Host "  devutils convert data.csv --to json --output result.json"
    Write-Host "  devutils convert data.txt --from csv --to json"
    Write-Host ""
    Write-Host "Options:"
    Write-Host "  --to <target>       Required target format: xlsx, csv, json, yaml, xml"
    Write-Host "  --from <source>     Optional source override: csv, xlsx, json, yaml, yml, xml"
    Write-Host "  --output <file>     Optional output path. Defaults beside input with target extension."
    Write-Host "  --force             Overwrite an existing output file."
    Write-Host "  --help              Show this help."
    Write-Host ""
    Write-Host "Supported conversions:"
    Write-Host "  csv  -> xlsx"
    Write-Host "  xlsx -> csv"
    Write-Host "  csv  -> json"
    Write-Host "  json -> csv"
    Write-Host "  json -> yaml"
    Write-Host "  yaml -> json"
    Write-Host "  xml  -> json"
    Write-Host "  json -> xml"
    Write-Host ""
    Write-Host "Optional dependencies:"
    Write-Host "  XLSX conversions require ImportExcel:"
    Write-Host "    Install-Module ImportExcel -Scope CurrentUser"
    Write-Host "  YAML conversions require powershell-yaml:"
    Write-Host "    Install-Module powershell-yaml -Scope CurrentUser"
    Write-Host ""
    Write-Host "Limitations:"
    Write-Host "  JSON -> CSV supports arrays of flat objects only."
    Write-Host "  CSV values are text; type information is not inferred."
    Write-Host "  XML/JSON mapping preserves elements and attributes as @attribute keys, but namespaces, ordering, comments, and some type information are not round-trip lossless."
}

Export-ModuleMember -Function Invoke-FileConversion

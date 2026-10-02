function Test-ImportExcelDependency {
    if (!(Get-Module -ListAvailable -Name ImportExcel)) {
        throw @"
ImportExcel PowerShell module is required for XLSX conversions.

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

function New-OutputFilePath {
    param(
        [Parameter(Mandatory = $true)][string]$InputFile,
        [Parameter(Mandatory = $true)][string]$TargetExtension,
        [string]$OutputFile
    )

    if (![string]::IsNullOrWhiteSpace($OutputFile)) {
        return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFile)
    }

    $Directory = Split-Path -Path $InputFile
    $FileName = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)
    return (Join-Path $Directory "$FileName$TargetExtension")
}

function Get-ConversionMap {
    return [ordered]@{
        ".csv->to-xlsx" = @{
            Source = ".csv"; TargetFormat = "to-xlsx"; TargetExtension = ".xlsx"; Description = ".csv        -> .xlsx"
            Converter = { param($Source, $Destination) Convert-CsvToXlsx -InputFile $Source -OutputFile $Destination }
        }
        ".xlsx->to-csv" = @{
            Source = ".xlsx"; TargetFormat = "to-csv"; TargetExtension = ".csv"; Description = ".xlsx       -> .csv"
            Converter = { param($Source, $Destination) Convert-XlsxToCsv -InputFile $Source -OutputFile $Destination }
        }
        ".csv->to-json" = @{
            Source = ".csv"; TargetFormat = "to-json"; TargetExtension = ".json"; Description = ".csv        -> .json"
            Converter = { param($Source, $Destination) Convert-CsvToJson -InputFile $Source -OutputFile $Destination }
        }
        ".json->to-csv" = @{
            Source = ".json"; TargetFormat = "to-csv"; TargetExtension = ".csv"; Description = ".json       -> .csv"
            Converter = { param($Source, $Destination) Convert-JsonToCsv -InputFile $Source -OutputFile $Destination }
        }
        ".json->to-yaml" = @{
            Source = ".json"; TargetFormat = "to-yaml"; TargetExtension = ".yaml"; Description = ".json       -> .yaml"
            Converter = { param($Source, $Destination) Convert-JsonToYaml -InputFile $Source -OutputFile $Destination }
        }
        ".yaml->to-json" = @{
            Source = ".yaml"; TargetFormat = "to-json"; TargetExtension = ".json"; Description = ".yaml/.yml  -> .json"
            Converter = { param($Source, $Destination) Convert-YamlToJson -InputFile $Source -OutputFile $Destination }
        }
        ".yml->to-json" = @{
            Source = ".yml"; TargetFormat = "to-json"; TargetExtension = ".json"; Description = ".yaml/.yml  -> .json"
            Converter = { param($Source, $Destination) Convert-YamlToJson -InputFile $Source -OutputFile $Destination }
        }
        ".xml->to-json" = @{
            Source = ".xml"; TargetFormat = "to-json"; TargetExtension = ".json"; Description = ".xml        -> .json"
            Converter = { param($Source, $Destination) Convert-XmlToJson -InputFile $Source -OutputFile $Destination }
        }
        ".json->to-xml" = @{
            Source = ".json"; TargetFormat = "to-xml"; TargetExtension = ".xml"; Description = ".json       -> .xml"
            Converter = { param($Source, $Destination) Convert-JsonToXml -InputFile $Source -OutputFile $Destination }
        }
        ".properties->to-yaml" = @{
            Source = ".properties"; TargetFormat = "to-yaml"; TargetExtension = ".yaml"; Description = ".properties -> .yaml"
            Converter = { param($Source, $Destination) Convert-PropertiesToYaml -InputFile $Source -OutputFile $Destination }
        }
        ".yaml->to-properties" = @{
            Source = ".yaml"; TargetFormat = "to-properties"; TargetExtension = ".properties"; Description = ".yaml/.yml  -> .properties"
            Converter = { param($Source, $Destination) Convert-YamlToProperties -InputFile $Source -OutputFile $Destination }
        }
        ".yml->to-properties" = @{
            Source = ".yml"; TargetFormat = "to-properties"; TargetExtension = ".properties"; Description = ".yaml/.yml  -> .properties"
            Converter = { param($Source, $Destination) Convert-YamlToProperties -InputFile $Source -OutputFile $Destination }
        }
    }
}

function Get-SupportedConversions {
    return (Get-ConversionMap).Values | ForEach-Object { $_["Description"] } | Select-Object -Unique
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

function ConvertTo-YamlScalar {
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return "null" }
    if ($Value -is [bool]) { return $Value.ToString().ToLowerInvariant() }
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [decimal] -or $Value -is [double]) {
        return $Value.ToString([System.Globalization.CultureInfo]::InvariantCulture)
    }

    $Text = [string]$Value
    if ($Text -eq "") { return "''" }

    if ($Text -match "[:#\[\]\{\},&\*!|>'`"%@]" -or $Text.Trim() -ne $Text -or $Text -match "^(true|false|null|[-+]?\d+(\.\d+)?)$") {
        return "'" + $Text.Replace("'", "''") + "'"
    }

    return $Text
}

function ConvertTo-SimpleYamlLines {
    param([AllowNull()]$Value, [int]$Indent = 0)

    $Padding = " " * $Indent
    $Value = ConvertTo-PlainObject $Value
    $Lines = @()

    if ($Value -is [System.Collections.IDictionary]) {
        foreach ($Key in $Value.Keys) {
            $Child = $Value[$Key]
            if ($Child -is [System.Collections.IDictionary] -or ($Child -is [System.Collections.IEnumerable] -and $Child -isnot [string])) {
                $Lines += "$Padding${Key}:"
                $Lines += ConvertTo-SimpleYamlLines -Value $Child -Indent ($Indent + 2)
            }
            else {
                $Lines += "$Padding${Key}: $(ConvertTo-YamlScalar $Child)"
            }
        }
        return $Lines
    }

    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        foreach ($Item in $Value) {
            if ($Item -is [System.Collections.IDictionary] -or ($Item -is [System.Collections.IEnumerable] -and $Item -isnot [string])) {
                $Lines += "$Padding-"
                $Lines += ConvertTo-SimpleYamlLines -Value $Item -Indent ($Indent + 2)
            }
            else {
                $Lines += "$Padding- $(ConvertTo-YamlScalar $Item)"
            }
        }
        return $Lines
    }

    return @("$Padding$(ConvertTo-YamlScalar $Value)")
}

function ConvertFrom-YamlScalar {
    param([string]$Value)

    $Value = $Value.Trim()
    if ($Value -eq "" -or $Value -eq "null" -or $Value -eq "~") { return $null }
    if ($Value -eq "true") { return $true }
    if ($Value -eq "false") { return $false }

    if (($Value.StartsWith("'") -and $Value.EndsWith("'")) -or ($Value.StartsWith('"') -and $Value.EndsWith('"'))) {
        return $Value.Substring(1, $Value.Length - 2).Replace("''", "'")
    }

    $Number = 0.0
    if ([double]::TryParse($Value, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$Number)) {
        return $Number
    }

    return $Value
}

function ConvertFrom-SimpleYaml {
    param([Parameter(Mandatory = $true)][string[]]$Lines)

    $Root = [ordered]@{}
    $Stack = @(@{ Indent = -1; Value = $Root })

    foreach ($Line in $Lines) {
        if ([string]::IsNullOrWhiteSpace($Line) -or $Line.TrimStart().StartsWith("#")) { continue }

        $Indent = $Line.Length - $Line.TrimStart().Length
        $Text = $Line.Trim()

        while ($Stack.Count -gt 1 -and $Indent -le $Stack[-1].Indent) {
            $Stack = $Stack[0..($Stack.Count - 2)]
        }

        $Parent = $Stack[-1].Value

        if ($Text.StartsWith("- ")) {
            throw "YAML list parsing is only supported inside named properties."
        }

        if ($Text -notmatch "^([^:]+):(.*)$") {
            throw "Unsupported YAML line: $Line"
        }

        $Key = $Matches[1].Trim()
        $RawValue = $Matches[2].Trim()

        if ($RawValue -eq "") {
            $Child = [ordered]@{}
            $Parent[$Key] = $Child
            $Stack += @{ Indent = $Indent; Value = $Child }
        }
        elseif ($RawValue.StartsWith("[") -and $RawValue.EndsWith("]")) {
            $Items = $RawValue.Trim("[", "]").Split(",") |
                Where-Object { $_.Trim() -ne "" } |
                ForEach-Object { ConvertFrom-YamlScalar $_ }
            $Parent[$Key] = @($Items)
        }
        else {
            $Parent[$Key] = ConvertFrom-YamlScalar $RawValue
        }
    }

    return $Root
}

function Read-JsonFile {
    param([Parameter(Mandatory = $true)][string]$InputFile)

    $Content = Get-Content -LiteralPath $InputFile -Raw
    if ([string]::IsNullOrWhiteSpace($Content)) {
        throw "JSON file contains no data."
    }

    return ($Content | ConvertFrom-Json)
}

function Read-PropertiesFile {
    param([Parameter(Mandatory = $true)][string]$InputFile)

    $Properties = [ordered]@{}

    foreach ($Line in Get-Content -LiteralPath $InputFile) {
        $Trimmed = $Line.Trim()
        if ($Trimmed -eq "" -or $Trimmed.StartsWith("#") -or $Trimmed.StartsWith("!")) { continue }

        if ($Trimmed -notmatch "^([^:=\s]+)\s*[:=]\s*(.*)$") {
            throw "Unsupported properties line: $Line"
        }

        $Properties[$Matches[1]] = $Matches[2]
    }

    if ($Properties.Count -eq 0) {
        throw "Properties file contains no data."
    }

    return $Properties
}

function ConvertTo-PropertiesLines {
    param([Parameter(Mandatory = $true)][System.Collections.IDictionary]$Value, [string]$Prefix = "")

    $Lines = @()

    foreach ($Key in $Value.Keys) {
        $Name = if ($Prefix) { "$Prefix.$Key" } else { $Key }
        $Child = $Value[$Key]

        if ($Child -is [System.Collections.IDictionary]) {
            $Lines += ConvertTo-PropertiesLines -Value $Child -Prefix $Name
        }
        elseif ($Child -is [System.Collections.IEnumerable] -and $Child -isnot [string]) {
            $Lines += "$Name=$($Child -join ',')"
        }
        else {
            $Lines += "$Name=$Child"
        }
    }

    return $Lines
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
        $ChildValue = Convert-XmlNodeToObject $Child
        $ChildName = $Child.LocalName

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

function Write-ConversionHeader {
    param([string]$Label, [string]$InputFile, [string]$OutputFile)

    Write-Host ""
    Write-Host $Label
    Write-Host "-----------------------------"
    Write-Host "Input : $InputFile"
    Write-Host "Output: $OutputFile"
    Write-Host ""
}

function Write-ConversionSuccess {
    param([string]$OutputFile)

    Write-Host "Conversion completed successfully."
    Write-Host ""
    Write-Host $OutputFile
}

function Convert-CsvToXlsx {
    param([string]$InputFile, [string]$OutputFile)

    Test-ImportExcelDependency
    Write-ConversionHeader -Label "CSV -> XLSX" -InputFile $InputFile -OutputFile $OutputFile

    try {
        $Data = Import-Csv -LiteralPath $InputFile
        if ($null -eq $Data -or $Data.Count -eq 0) { throw "CSV file contains no data." }

        $Data | Export-Excel -Path $OutputFile -WorksheetName "Data" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
        Write-ConversionSuccess $OutputFile
    }
    catch {
        throw "CSV to XLSX conversion failed: $($_.Exception.Message)"
    }
}

function Convert-XlsxToCsv {
    param([string]$InputFile, [string]$OutputFile)

    Test-ImportExcelDependency
    Write-ConversionHeader -Label "XLSX -> CSV" -InputFile $InputFile -OutputFile $OutputFile

    try {
        $Data = Import-Excel -Path $InputFile
        if ($null -eq $Data -or $Data.Count -eq 0) { throw "XLSX file contains no data." }

        $Data | Export-Csv -LiteralPath $OutputFile -NoTypeInformation
        Write-ConversionSuccess $OutputFile
    }
    catch {
        throw "XLSX to CSV conversion failed: $($_.Exception.Message)"
    }
}

function Convert-CsvToJson {
    param([string]$InputFile, [string]$OutputFile)

    Write-ConversionHeader -Label "CSV -> JSON" -InputFile $InputFile -OutputFile $OutputFile

    try {
        $Data = Import-Csv -LiteralPath $InputFile
        if ($null -eq $Data -or $Data.Count -eq 0) { throw "CSV file contains no data." }

        $Data | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $OutputFile
        Write-ConversionSuccess $OutputFile
    }
    catch {
        throw "CSV to JSON conversion failed: $($_.Exception.Message)"
    }
}

function Convert-JsonToCsv {
    param([string]$InputFile, [string]$OutputFile)

    Write-ConversionHeader -Label "JSON -> CSV" -InputFile $InputFile -OutputFile $OutputFile

    try {
        $Data = Read-JsonFile $InputFile
        if ($Data -isnot [System.Collections.IEnumerable] -or $Data -is [string]) { $Data = @($Data) }
        if ($Data.Count -eq 0) { throw "JSON file contains no data." }

        $Data | Export-Csv -LiteralPath $OutputFile -NoTypeInformation
        Write-ConversionSuccess $OutputFile
    }
    catch {
        throw "JSON to CSV conversion failed: $($_.Exception.Message)"
    }
}

function Convert-JsonToYaml {
    param([string]$InputFile, [string]$OutputFile)

    Write-ConversionHeader -Label "JSON -> YAML" -InputFile $InputFile -OutputFile $OutputFile

    try {
        $Data = Read-JsonFile $InputFile
        ConvertTo-SimpleYamlLines $Data | Set-Content -LiteralPath $OutputFile
        Write-ConversionSuccess $OutputFile
    }
    catch {
        throw "JSON to YAML conversion failed: $($_.Exception.Message)"
    }
}

function Convert-YamlToJson {
    param([string]$InputFile, [string]$OutputFile)

    Write-ConversionHeader -Label "YAML -> JSON" -InputFile $InputFile -OutputFile $OutputFile

    try {
        $Data = ConvertFrom-SimpleYaml -Lines (Get-Content -LiteralPath $InputFile)
        if ($Data.Count -eq 0) { throw "YAML file contains no data." }

        $Data | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $OutputFile
        Write-ConversionSuccess $OutputFile
    }
    catch {
        throw "YAML to JSON conversion failed: $($_.Exception.Message)"
    }
}

function Convert-XmlToJson {
    param([string]$InputFile, [string]$OutputFile)

    Write-ConversionHeader -Label "XML -> JSON" -InputFile $InputFile -OutputFile $OutputFile

    try {
        [xml]$Xml = Get-Content -LiteralPath $InputFile -Raw
        $RootName = $Xml.DocumentElement.LocalName
        $Data = [ordered]@{}
        $Data[$RootName] = Convert-XmlNodeToObject $Xml.DocumentElement
        $Data | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $OutputFile
        Write-ConversionSuccess $OutputFile
    }
    catch {
        throw "XML to JSON conversion failed: $($_.Exception.Message)"
    }
}

function Convert-JsonToXml {
    param([string]$InputFile, [string]$OutputFile)

    Write-ConversionHeader -Label "JSON -> XML" -InputFile $InputFile -OutputFile $OutputFile

    try {
        $Data = ConvertTo-PlainObject (Read-JsonFile $InputFile)
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
        Write-ConversionSuccess $OutputFile
    }
    catch {
        throw "JSON to XML conversion failed: $($_.Exception.Message)"
    }
}

function Convert-PropertiesToYaml {
    param([string]$InputFile, [string]$OutputFile)

    Write-ConversionHeader -Label "Properties -> YAML" -InputFile $InputFile -OutputFile $OutputFile

    try {
        $Data = Read-PropertiesFile $InputFile
        ConvertTo-SimpleYamlLines $Data | Set-Content -LiteralPath $OutputFile
        Write-ConversionSuccess $OutputFile
    }
    catch {
        throw "Properties to YAML conversion failed: $($_.Exception.Message)"
    }
}

function Convert-YamlToProperties {
    param([string]$InputFile, [string]$OutputFile)

    Write-ConversionHeader -Label "YAML -> Properties" -InputFile $InputFile -OutputFile $OutputFile

    try {
        $Data = ConvertFrom-SimpleYaml -Lines (Get-Content -LiteralPath $InputFile)
        if ($Data.Count -eq 0) { throw "YAML file contains no data." }

        ConvertTo-PropertiesLines -Value $Data | Set-Content -LiteralPath $OutputFile
        Write-ConversionSuccess $OutputFile
    }
    catch {
        throw "YAML to Properties conversion failed: $($_.Exception.Message)"
    }
}

function Invoke-FileConversion {
    param(
        [Parameter(Position = 0)][string]$TargetFormat,
        [Parameter(Position = 1)][string]$InputFile,
        [Parameter(Position = 2)][string]$OutputFile
    )

    if ([string]::IsNullOrWhiteSpace($TargetFormat)) {
        throw "Missing target format. Use: devutils convert help"
    }

    if ([string]::IsNullOrWhiteSpace($InputFile)) {
        throw "Missing input file. Use: devutils convert $TargetFormat <input-file> [output-file]"
    }

    $TargetFormat = $TargetFormat.ToLowerInvariant()
    $Conversions = Get-ConversionMap
    $SupportedTargetFormats = $Conversions.Values | ForEach-Object { $_["TargetFormat"] } | Select-Object -Unique

    if ($TargetFormat -notin $SupportedTargetFormats) {
        throw "Unsupported target format: $TargetFormat"
    }

    $InputFile = Resolve-InputFilePath -InputFile $InputFile
    $SourceExtension = [System.IO.Path]::GetExtension($InputFile).ToLowerInvariant()
    $SupportedSourceExtensions = $Conversions.Values | ForEach-Object { $_["Source"] } | Select-Object -Unique

    if ($SourceExtension -notin $SupportedSourceExtensions) {
        throw "Unsupported source extension: $SourceExtension"
    }

    $Route = "$SourceExtension->$TargetFormat"

    if (!$Conversions.Contains($Route)) {
        throw "Unsupported conversion: $Route"
    }

    $Conversion = $Conversions[$Route]
    $OutputFile = New-OutputFilePath -InputFile $InputFile -TargetExtension $Conversion["TargetExtension"] -OutputFile $OutputFile

    try {
        & $Conversion["Converter"] $InputFile $OutputFile
    }
    catch {
        throw "File conversion failed: $($_.Exception.Message)"
    }
}

Export-ModuleMember -Function Invoke-FileConversion

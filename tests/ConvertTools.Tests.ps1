$Script:RepoRoot = Split-Path -Parent $PSScriptRoot
$Script:Devutils = Join-Path $Script:RepoRoot "devutils.ps1"
$Script:Fixtures = Join-Path $PSScriptRoot "fixtures"

function Invoke-DevutilsCli {
    param([string[]]$Arguments)

    $Output = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $Script:Devutils @Arguments 2>&1

    [pscustomobject]@{
        ExitCode = $LASTEXITCODE
        Output = ($Output | Out-String)
    }
}

Describe "ConvertTools CLI" {

    It "converts CSV to JSON using source inference" {
        $InputFile = Join-Path $Script:Fixtures "people.csv"
        $OutputFile = Join-Path $TestDrive "people.json"

        $Result = Invoke-DevutilsCli @("convert", $InputFile, "--to", "json", "--output", $OutputFile)

        $Result.ExitCode | Should Be 0
        Test-Path -LiteralPath $OutputFile | Should Be $true
        (Get-Content -LiteralPath $OutputFile -Raw) | Should Match "caf"
    }

    It "uses explicit source format override for ambiguous TXT input" {
        $InputFile = Join-Path $TestDrive "people.txt"
        Copy-Item -LiteralPath (Join-Path $Script:Fixtures "people.csv") -Destination $InputFile
        $OutputFile = Join-Path $TestDrive "people-from-txt.json"

        $Result = Invoke-DevutilsCli @("convert", $InputFile, "--from", "csv", "--to", "json", "--output", $OutputFile)

        $Result.ExitCode | Should Be 0
        Test-Path -LiteralPath $OutputFile | Should Be $true
    }

    It "rejects TXT input without explicit source format" {
        $InputFile = Join-Path $TestDrive "ambiguous.txt"
        Copy-Item -LiteralPath (Join-Path $Script:Fixtures "people.csv") -Destination $InputFile

        $Result = Invoke-DevutilsCli @("convert", $InputFile, "--to", "json")

        $Result.ExitCode | Should Be 1
        $Result.Output | Should Match "Ambiguous source extension"
    }

    It "requires --to" {
        $InputFile = Join-Path $Script:Fixtures "people.csv"

        $Result = Invoke-DevutilsCli @("convert", $InputFile)

        $Result.ExitCode | Should Be 1
        $Result.Output | Should Match "Missing required option"
    }

    It "requires an input file" {
        $Result = Invoke-DevutilsCli @("convert")

        $Result.ExitCode | Should Be 1
        $Result.Output | Should Match "Missing input file"
    }

    It "rejects unknown options" {
        $InputFile = Join-Path $Script:Fixtures "people.csv"

        $Result = Invoke-DevutilsCli @("convert", $InputFile, "--to", "json", "--bogus")

        $Result.ExitCode | Should Be 1
        $Result.Output | Should Match "Unknown option"
    }

    It "protects existing output files unless --force is supplied" {
        $InputFile = Join-Path $Script:Fixtures "people.csv"
        $OutputFile = Join-Path $TestDrive "existing.json"
        Set-Content -LiteralPath $OutputFile -Value "already here"

        $Blocked = Invoke-DevutilsCli @("convert", $InputFile, "--to", "json", "--output", $OutputFile)
        $Forced = Invoke-DevutilsCli @("convert", $InputFile, "--to", "json", "--output", $OutputFile, "--force")

        $Blocked.ExitCode | Should Be 1
        $Blocked.Output | Should Match "already exists"
        $Forced.ExitCode | Should Be 0
    }

    It "converts JSON arrays of flat objects to CSV" {
        $InputFile = Join-Path $Script:Fixtures "people.json"
        $OutputFile = Join-Path $TestDrive "people.csv"

        $Result = Invoke-DevutilsCli @("convert", $InputFile, "--to", "csv", "--output", $OutputFile)

        $Result.ExitCode | Should Be 0
        (Get-Content -LiteralPath $OutputFile -Raw) | Should Match '"Name","Role"'
    }

    It "rejects nested JSON when converting to CSV" {
        $InputFile = Join-Path $Script:Fixtures "nested.json"
        $OutputFile = Join-Path $TestDrive "nested.csv"

        $Result = Invoke-DevutilsCli @("convert", $InputFile, "--to", "csv", "--output", $OutputFile)

        $Result.ExitCode | Should Be 1
        $Result.Output | Should Match "nested property"
    }

    It "converts XML to JSON with attributes represented as @ keys" {
        $InputFile = Join-Path $Script:Fixtures "person.xml"
        $OutputFile = Join-Path $TestDrive "person.json"

        $Result = Invoke-DevutilsCli @("convert", $InputFile, "--to", "json", "--output", $OutputFile)

        $Result.ExitCode | Should Be 0
        $Json = Get-Content -LiteralPath $OutputFile -Raw
        $Json | Should Match '"person"'
        $Json | Should Match '"@id"'
    }

    It "handles paths containing spaces" {
        $Directory = Join-Path $TestDrive "folder with spaces"
        New-Item -ItemType Directory -Path $Directory | Out-Null
        $InputFile = Join-Path $Directory "people data.csv"
        $OutputFile = Join-Path $Directory "people data.json"
        Copy-Item -LiteralPath (Join-Path $Script:Fixtures "people.csv") -Destination $InputFile

        $Result = Invoke-DevutilsCli @("convert", $InputFile, "--to", "json", "--output", $OutputFile)

        $Result.ExitCode | Should Be 0
        Test-Path -LiteralPath $OutputFile | Should Be $true
    }

    It "reports missing YAML dependency when powershell-yaml is unavailable" {
        if (Get-Module -ListAvailable -Name powershell-yaml) {
            Set-TestInconclusive -Message "powershell-yaml is installed on this machine."
        }

        $InputFile = Join-Path $Script:Fixtures "config.yaml"
        $OutputFile = Join-Path $TestDrive "config.json"

        $Result = Invoke-DevutilsCli @("convert", $InputFile, "--to", "json", "--output", $OutputFile)

        $Result.ExitCode | Should Be 1
        $Result.Output | Should Match "powershell-yaml"
    }
}

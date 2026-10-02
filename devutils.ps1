param(
    [Parameter(Position = 0)]
    [string]$Category,

    [Parameter(Position = 1)]
    [string]$Command,

    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Arguments
)

$ErrorActionPreference = "Stop"

$RootPath = $PSScriptRoot
$ModulePath = Join-Path $RootPath "modules"

function Show-Help {
    Write-Host ""
    Write-Host "DevUtils-PS - Developer Productivity Toolkit"
    Write-Host ""
    Write-Host "Usage:"
    Write-Host "  devutils <category> <command> [arguments]"
    Write-Host "  devutils <category> help"
    Write-Host ""
    Write-Host "Commands:"
    Write-Host "  log     search          Search text in log file"
    Write-Host "  json    format          Pretty-print JSON"
    Write-Host "  config  compare         Compare configuration files"
    Write-Host "  net     check           Check URL health"
    Write-Host "  git     compare         Compare Git branches"
    Write-Host "  convert to-xlsx         Convert CSV to XLSX"
    Write-Host "  convert to-csv          Convert XLSX to CSV"
    Write-Host "  convert to-json         Convert CSV, YAML, or XML to JSON"
    Write-Host "  convert to-yaml         Convert JSON or properties to YAML"
    Write-Host "  convert to-xml          Convert JSON to XML"
    Write-Host "  convert to-properties   Convert YAML to properties"
    Write-Host "  excel   date-format     Convert date formats in an XLSX column"
    Write-Host ""
    Write-Host "Examples:"
    Write-Host '  devutils log search app.log "ERROR"'
    Write-Host '  devutils json format response.json'
    Write-Host '  devutils config compare uat.properties prod.properties'
    Write-Host '  devutils net check urls.txt'
    Write-Host '  devutils git compare dev testing'
    Write-Host '  devutils convert to-xlsx input.csv'
    Write-Host '  devutils convert to-csv input.xlsx'
    Write-Host '  devutils convert to-json input.csv'
    Write-Host '  devutils convert to-yaml input.properties'
    Write-Host '  devutils excel date-format report.xlsx Date dd/MM/yyyy yyyy-MM-dd'
    Write-Host ""
    Write-Host "  csv     to-xlsx         Convert CSV to XLSX"
    Write-Host '  devutils csv to-xlsx input.csv'
}

function Test-HelpCommand {
    return ([string]::IsNullOrWhiteSpace($Command) -or $Command.ToLower() -eq "help")
}

function Show-LogHelp {
    Write-Host ""
    Write-Host "Log commands:"
    Write-Host "  devutils log search <log-file> <search-text>"
}

function Show-JsonHelp {
    Write-Host ""
    Write-Host "JSON commands:"
    Write-Host "  devutils json format <json-file>"
}

function Show-ConfigHelp {
    Write-Host ""
    Write-Host "Config commands:"
    Write-Host "  devutils config compare <left-file> <right-file>"
}

function Show-NetworkHelp {
    Write-Host ""
    Write-Host "Network commands:"
    Write-Host "  devutils net check <urls-file>"
}

function Show-GitHelp {
    Write-Host ""
    Write-Host "Git commands:"
    Write-Host "  devutils git compare <source-branch> <target-branch>"
}

function Show-ConvertHelp {
    Write-Host ""
    Write-Host "Conversion commands:"
    Write-Host "  devutils convert to-xlsx <input.csv> [output.xlsx]"
    Write-Host "  devutils convert to-csv  <input.xlsx> [output.csv]"
    Write-Host "  devutils convert to-json <input.csv|input.yaml|input.yml|input.xml> [output.json]"
    Write-Host "  devutils convert to-yaml <input.json|input.properties> [output.yaml]"
    Write-Host "  devutils convert to-xml <input.json> [output.xml]"
    Write-Host "  devutils convert to-properties <input.yaml|input.yml> [output.properties]"
    Write-Host ""
    Write-Host "Supported conversions:"
    Write-Host "  .csv        -> .xlsx"
    Write-Host "  .xlsx       -> .csv"
    Write-Host "  .csv        -> .json"
    Write-Host "  .json       -> .csv"
    Write-Host "  .json       -> .yaml"
    Write-Host "  .yaml/.yml  -> .json"
    Write-Host "  .xml        -> .json"
    Write-Host "  .json       -> .xml"
    Write-Host "  .properties -> .yaml"
    Write-Host "  .yaml/.yml  -> .properties"
}

function Show-CsvHelp {
    Write-Host ""
    Write-Host "CSV commands:"
    Write-Host "  devutils csv to-xlsx <input.csv> [output.xlsx]"
}

function Show-ExcelHelp {
    Write-Host ""
    Write-Host "Excel commands:"
    Write-Host "  devutils excel date-format --input-file <input.xlsx> --column <column> --from-format <format> --to-format <format> [--output-file <output.xlsx>] [--sheet-name <worksheet>]"
    Write-Host ""
    Write-Host "Examples:"
    Write-Host "  devutils excel date-format --input-file report.xlsx --column Date --from-format dd/MM/yyyy --to-format yyyy-MM-dd"
    Write-Host "  devutils excel date-format --input-file report.xlsx --column C --from-format MM-dd-yyyy --to-format dd-MMM-yyyy --output-file fixed.xlsx --sheet-name Sheet1"
}

if ([string]::IsNullOrWhiteSpace($Category)) {
    Show-Help
    exit 0
}

try {

    switch ($Category.ToLower()) {

        "log" {
            Import-Module (Join-Path $ModulePath "LogTools.psm1") -Force

            if (Test-HelpCommand) {
                Show-LogHelp
                exit 0
            }

            switch ($Command.ToLower()) {

                "search" {
                    Invoke-LogSearch @Arguments
                }

                default {
                    Write-Error "Unknown log command: $Command"
                }
            }
        }

        "json" {
            Import-Module (Join-Path $ModulePath "JsonTools.psm1") -Force

            if (Test-HelpCommand) {
                Show-JsonHelp
                exit 0
            }

            switch ($Command.ToLower()) {

                "format" {
                    Format-JsonFile @Arguments
                }

                default {
                    Write-Error "Unknown json command: $Command"
                }
            }
        }

        "config" {
            Import-Module (Join-Path $ModulePath "ConfigTools.psm1") -Force

            if (Test-HelpCommand) {
                Show-ConfigHelp
                exit 0
            }

            switch ($Command.ToLower()) {

                "compare" {
                    Compare-ConfigFile @Arguments
                }

                default {
                    Write-Error "Unknown config command: $Command"
                }
            }
        }

        "net" {
            Import-Module (Join-Path $ModulePath "NetworkTools.psm1") -Force

            if (Test-HelpCommand) {
                Show-NetworkHelp
                exit 0
            }

            switch ($Command.ToLower()) {

                "check" {
                    Test-UrlList @Arguments
                }

                default {
                    Write-Error "Unknown net command: $Command"
                }
            }
        }

        "git" {
            Import-Module (Join-Path $ModulePath "GitTools.psm1") -Force

            if (Test-HelpCommand) {
                Show-GitHelp
                exit 0
            }

            switch ($Command.ToLower()) {

                "compare" {
                    Compare-GitBranches @Arguments
                }

                default {
                    Write-Error "Unknown git command: $Command"
                }
            }
        }

        "convert" {
            Import-Module (Join-Path $ModulePath "ConvertTools.psm1") -Force

            if (Test-HelpCommand) {
                Show-ConvertHelp
                exit 0
            }

            Invoke-FileConversion $Command @Arguments
        }

        "csv" {

            Import-Module (Join-Path $ModulePath "CsvTools.psm1") -Force

            if (Test-HelpCommand) {
                Show-CsvHelp
                exit 0
            }

            switch ($Command.ToLower()) {

                "to-xlsx" {
                    Convert-CsvToXlsx @Arguments
                }

                default {
                    Write-Error "Unknown CSV command: $Command"
                }
            }
        }

        "excel" {
            Import-Module (Join-Path $ModulePath "ExcelTools.psm1") -Force

            if (Test-HelpCommand) {
                Show-ExcelHelp
                exit 0
            }

            switch ($Command.ToLower()) {

                "date-format" {
                    Convert-ExcelColumnDateFormat @Arguments
                }

                default {
                    Write-Error "Unknown excel command: $Command"
                }
            }
        }

        "help" {
            Show-Help
        }

        default {
            Write-Host "Unknown category: $Category"
            Show-Help
            exit 1
        }
    }
}
catch {
    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

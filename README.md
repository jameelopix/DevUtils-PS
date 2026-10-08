# DevUtils-PS

Portable PowerShell developer utilities for restricted corporate environments.

## Usage

```powershell
devutils <category> <command> [arguments]
devutils <category> help
```

## Generic File Conversion

```powershell
devutils convert <input> --to <target>
devutils convert <input> --from <source> --to <target>
```

Examples:

```powershell
devutils convert data.csv --to xlsx
devutils convert data.xlsx --to csv
devutils convert data.csv --to json
devutils convert data.json --to csv
devutils convert data.json --to yaml
devutils convert data.yaml --to json
devutils convert data.xml --to json
devutils convert data.json --to xml
devutils convert data.txt --from csv --to json
devutils convert data.csv --to json --output result.json
```

Options:

```text
--to <target>       Required target format: xlsx, csv, json, yaml, xml
--from <source>     Optional source override: csv, xlsx, json, yaml, yml, xml
--output <file>     Optional output path. Defaults beside input with target extension.
--force             Overwrite an existing output file.
--help              Show conversion help.
```

Supported conversions:

| Source | Target |
|---|---|
| CSV | XLSX |
| XLSX | CSV |
| CSV | JSON |
| JSON | CSV |
| JSON | YAML |
| YAML/YML | JSON |
| XML | JSON |
| JSON | XML |

Source format is inferred from the input extension unless `--from` is supplied. Ambiguous extensions such as `.txt` are not guessed; use `--from`.

Default output uses the input filename with the target extension. Existing files are not overwritten unless `--force` is supplied.

### Optional Dependencies

XLSX conversions require ImportExcel:

```powershell
Install-Module ImportExcel -Scope CurrentUser
```

YAML conversions require powershell-yaml:

```powershell
Install-Module powershell-yaml -Scope CurrentUser
```

DevUtils-PS does not install dependencies automatically and does not require administrator privileges.

### Conversion Notes

CSV parsing and writing use PowerShell CSV cmdlets, so quoted values and headers are handled by the platform. CSV values are text; type information is not inferred.

JSON to CSV supports arrays of flat objects only. Nested JSON is rejected with a clear error instead of being silently flattened.

XML to JSON maps attributes to keys prefixed with `@`. JSON to XML maps `@name` keys back to attributes. XML comments, namespace details, ordering, and type information are not guaranteed to round-trip losslessly.

## Excel Utilities

Convert the date display format for one XLSX column:

```powershell
devutils excel date-format --input-file report.xlsx --column Date --from-format dd/MM/yyyy --to-format yyyy-MM-dd
devutils excel date-format --input-file report.xlsx --column C --from-format MM-dd-yyyy --to-format dd-MMM-yyyy --output-file fixed.xlsx --sheet-name Sheet1
```

The column can be a header name or an Excel column letter. Cells may contain Excel date values or strings matching `--from-format`.

Excel utilities require ImportExcel.

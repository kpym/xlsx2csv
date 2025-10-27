# xlsx2csv
Simple script for converting xlsx files to csv files commandline. 

The code of this tool is heavily based on [tealeg/xlsx](https://github.com/tealeg/xlsx2csv)

## Usage
```
> xlsx2csv -h
xlsx2csv (version: 0.1.2) convert the given xlsx file to a csv.

Usage:
        xlsx2csv [flags] <xlsx-to-be-read>
  -d string
        Delimiter to use between fields (default ",")
  -i int
        Index of sheet to convert, zero based. (default -1)
  -o string
        filename to output to.

Defaults :
- If -i is not given or negative or is negative, all sheets are converted.       
- If -o is not given, output filename is derived from input filename by replacing .xlsx with .csv
- If -o is "stdout", output is written to standard output
- If multiple sheets are converted, the sheet index is appended to the output filename.
  If the output filename has a %d, it is replaced with the sheet index,
  if not, the index is added before the extension.
- If -d is not given, comma (,) is used as the delimiter

Examples:
- Convert all sheets in input.xlsx to CSV files named input.0.csv, input.1.csv, etc:
> xlsx2csv input.xlsx
- Convert only the second sheet (index 1) in input.xlsx to output.csv using semicolon as delimiter:
> xlsx2csv -i 1 -o output.csv -d ';' input.xlsx
- Convert the first sheet to stdout:
> xlsx2csv -o stdout input.xlsx
```

## Installation

Dowload it from the [releases page](https://github.com/kpym/xlsx2csv/releases) and put it in your path.
Or build it yourself:

```bash
go install github.com/kpym/xlsx2csv@latest
```

## License

[LICENSE](LICENSE)
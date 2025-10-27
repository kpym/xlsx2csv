// Copyright 2011-2015, The xlsx2csv Authors.
// All rights reserved.
// For details, see the LICENSE file.

package main

import (
	"encoding/csv"
	"errors"
	"flag"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/tealeg/xlsx/v3"
)

var version = "dev"

// csvName generate the name of the CSV file.
// It takes a pattern and sheet index to produce the filename.
// If the pattern hase no extension, .csv is appended.
// If the pattern has a %d, it is replaced with the sheet index.
// If the pattern hase no %d the index is added before the extension.
// If the sheet index is negative, no index is added (%d is removed).
// If the pattern is empty, an empty string is returned.
func csvName(pattern string, sheetIndex int) string {
	if pattern == "" {
		return ""
	}
	ext := filepath.Ext(pattern)
	base := pattern
	if ext == "" {
		ext = ".csv"
	} else {
		base = strings.TrimSuffix(pattern, ext)
	}
	if sheetIndex >= 0 {
		indexStr := fmt.Sprintf("%d", sheetIndex)
		if strings.Contains(base, "%d") {
			base = strings.Replace(base, "%d", indexStr, 1)
		} else {
			base = fmt.Sprintf("%s.%s", base, indexStr)
		}
	} else {
		base = strings.Replace(base, "%d", "", 1)
	}
	return base + ext
}

func generateCSVFromXLSXFile(xlFile *xlsx.File, sheetIndex int, csvOpts csvOptSetter, outName string) error {
	// determine writer: stdout if outName is empty, otherwise create the file
	var w io.Writer = os.Stdout
	var f *os.File
	if outName != "" {
		var err error
		f, err = os.Create(outName)
		if err != nil {
			return err
		}
		// ensure we close the file we created
		defer func() {
			_ = f.Close()
		}()
		w = f
	}

	// xlFile is opened by the caller; do not open it here.
	sheetLen := len(xlFile.Sheets)
	switch {
	case sheetLen == 0:
		return errors.New("This XLSX file contains no sheets.")
	case sheetIndex >= sheetLen:
		return fmt.Errorf("No sheet %d available, please select a sheet between 0 and %d\n", sheetIndex, sheetLen-1)
	}
	cw := csv.NewWriter(w)
	if csvOpts != nil {
		csvOpts(cw)
	}
	sheet := xlFile.Sheets[sheetIndex]
	var vals []string
	err := sheet.ForEachRow(func(row *xlsx.Row) error {
		if row != nil {
			vals = vals[:0]
			err := row.ForEachCell(func(cell *xlsx.Cell) error {
				str, err := cell.FormattedValue()
				if err != nil {
					return err
				}
				vals = append(vals, str)
				return nil
			})
			if err != nil {
				return err
			}
		}
		cw.Write(vals)
		return nil
	})
	if err != nil {
		return err
	}
	cw.Flush()
	return cw.Error()
}

type csvOptSetter func(*csv.Writer)

func main() {
	var (
		outFile    = flag.String("o", "", "filename to output to.")
		sheetIndex = flag.Int("i", -1, "Index of sheet to convert, zero based.")
		delimiter  = flag.String("d", ",", "Delimiter to use between fields")
	)
	flag.Usage = func() {
		exe := filepath.Base(os.Args[0])
		// strip .exe extension (if on Windows)
		exe = strings.TrimSuffix(exe, ".exe")
		fmt.Fprintf(os.Stderr, `%s (version: %s)
	dumps the given xlsx file's chosen sheet as a CSV,
	with the specified delimiter, into the specified output.

Usage:
	%s [flags] <xlsx-to-be-read>
`, exe, version, exe)
		flag.PrintDefaults()
		fmt.Fprintf(os.Stderr, `
Defaults :
- If -i is not given or negative or is negative, all sheets are converted.
- If -o is not given, output filename is derived from input filename by replacing .xlsx with .csv
- If -o is "stdout", output is written to standard output
- If multiple sheets are converted, the sheet index is appended to the output filename.
  If the output filename has a %%d, it is replaced with the sheet index, 
  if not, the index is added before the extension.
- If -d is not given, comma (,) is used as the delimiter

Examples:
- Convert all sheets in input.xlsx to CSV files named input.0.csv, input.1.csv, etc:
> %s input.xlsx
- Convert only the second sheet (index 1) in input.xlsx to output.csv using semicolon as delimiter:
> %s -i 1 -o output.csv -d ';' input.xlsx
- Convert the first sheet to stdout:
> %s -o stdout input.xlsx
`, exe, exe)
	}

	flag.Parse()
	if flag.NArg() != 1 {
		flag.Usage()
		os.Exit(1)
	}

	// open the xlsx file here and pass xlFile into generateCSVFromXLSXFile
	xlFile, err := xlsx.OpenFile(flag.Arg(0))
	if err != nil {
		log.Fatal(err)
	}

	// preserve previous behavior: treat "-" as stdout by passing empty outName
	outName := *outFile
	if outName == "" {
		outName = strings.TrimSuffix(flag.Arg(0), ".xlsx") + ".csv"
	}
	if outName == "stdout" {
		outName = ""
	}
	if outName == "" && *sheetIndex < 0 {
		*sheetIndex = 0 // default to first sheet when writing to stdout
	}
	// get the number of sheets in the file
	sheetLen := len(xlFile.Sheets)

	// validate the sheet index
	if sheetLen == 0 {
		log.Fatalf("This XLSX file contains no sheets.\n")
	}
	if *sheetIndex >= sheetLen {
		log.Fatalf("No sheet %d available, please select a sheet between 0 and %d. Or -1 to convert all sheets.\n", *sheetIndex, sheetLen-1)
	}

	// determine the range of sheets to convert
	first, last := 0, sheetLen-1
	if *sheetIndex >= 0 {
		first, last = *sheetIndex, *sheetIndex
	}

	outNameI := func(i int) string { return csvName(outName, i) }
	if first == last {
		outNameI = func(i int) string { return csvName(outName, -1) }
	}

	// create the delimiter option
	delimiterOption := func(cw *csv.Writer) { cw.Comma = ([]rune(*delimiter))[0] }

	// loop over the sheets to convert
	for i := first; i <= last; i++ {
		if err := generateCSVFromXLSXFile(
			xlFile,
			i,
			delimiterOption,
			outNameI(i),
		); err != nil {
			log.Fatal(err)
		}
	}
}

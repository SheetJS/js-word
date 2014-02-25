# xls

Currently a parser for XLS (BIFF8) and XML (2003/2004) files.  Cleanroom implementation from the Microsoft Open Specifications.

## Installation

In node:

    npm install xlsjs

In the browser:

    <script src="xls.js"></script>

## Usage

Simple usage (gets the value of cell A1 of sheet Sheet1):

    var XLS = require('xlsjs');
    var xls = XLS.readFile('test.xls');
    var Sheet1A1 = xls.Sheets['Sheet1']['A1'].v;

The node version installs a binary `xls2csv` which can read XLS files and output the contents in various formats.  The source is available at `xls2csv.njs` in the bin directory.

See <http://oss.sheetjs.com/js-xls/> for a browser example.

Some helper functions in `XLS.utils` generate different views of the sheets:

- `XLS.utils.sheet_to_csv` generates CSV
- `XLS.utils.sheet_to_row_object_array` interprets sheets as tables with a header column and generates an array of objects
- `XLS.utils.get_formulae` generates a list of formulae

For more details:

- `bin/xls2csv.njs` is a tool for node
- `index.html` is the live demo
- `bits/80_xls.js` contains the logic for generating CSV and JSON from sheets

## Cell Object Description

`.SheetNames` is an ordered list of the sheets in the workbook

`.Sheets[sheetname]` returns a data structure representing the sheet.  Each key
that does not start with `!` corresponds to a cell (using `A-1` notation).

`.Sheets[sheetname][address]` returns the specified cell:

- `.v` : the raw value of the cell
- `.w` : the formatted text of the cell (if applicable)
- `.t` : the type of the cell (constrained to the enumeration `ST_CellType` as documented in page 4215 of ISO/IEC 29500-1:2012(E) )
- `.f` : the formula of the cell (if applicable)
- `.z` : the number format string associated with the cell (if requested)

For dates, `.v` holds the raw date code from the sheet and `.w` holds the text

## Options

The exported `read` and `readFile` functions accept an options argument:

| Option Name | Default | Description |
| :---------- | ------: | :---------- |
| cellFormula | true    | Save formulae to the .f field ** |
| cellNF      | false   | Save number format string to the .z field |
| sheetRows   | 0       | If >0, read the first `sheetRows` rows ** |
| bookFiles   | false   | If true, add raw files to book object ** |
| bookProps   | false   | If true, only parse enough to get book metadata ** |
| bookSheets  | false   | If true, only parse enough to get the sheet names |

- Even if `cellNF` is false, formatted text (.w) will be generated
- In some cases, sheets may be parsed even if `bookSheets` is false.
- `bookSheets` and `bookProps` combine to give both sets of information
- `bookFiles` adds a `cfb` object
- `sheetRows-1` rows will be generated when looking at the JSON object output
  (since the header row is counted as a row when parsing the data)

## Other Notes

`CFB` refers to the Microsoft Compound File Binary Format, a container format for XLS as well as DOC and other pre-OOXML data formats.

The mechanism is split into a CFB parser (which scans through the file and produces concrete data chunks) and a Workbook parser (which does excel-specific parsing).  XML files are not processed by the CFB parser.

## Test Files

Test files are housed in [another repo](https://github.com/SheetJS/test_files).

Running `make init` will refresh the `test_files` submodule and get the files.

Tests utilize the mocha testing framework.  Travis-CI and Sauce Labs links:

 - <https://travis-ci.org/SheetJS/js-xls> for XLS module in node
 - <https://travis-ci.org/SheetJS/SheetJS.github.io> for XLS* modules
 - <https://saucelabs.com/u/sheetjs> for XLS* modules using Sauce Labs

## Contributing

Due to the precarious nature of the Open Specifications Promise, it is very important to ensure code is cleanroom.  Consult CONTRIBUTING.md

## XLSX/XLSM/XLSB Support

XLSX/XLSM/XLSB is available in [js-xlsx](https://github.com/SheetJS/js-xlsx).

## License

Please consult the attached LICENSE file for details.  All rights not explicitly granted by the Apache 2.0 license are reserved by the Original Author.

It is the opinion of the Original Author that this code conforms to the terms of the Microsoft Open Specifications Promise, falling under the same terms as OpenOffice (which is governed by the Apache License v2).  Given the vagaries of the promise, the Original Author makes no legal claim that in fact end users are protected from future actions.  It is highly recommended that, for commercial uses, you consult a lawyer before proceeding.

## References

OSP-covered specifications:
 - [MS-CFB]: Compound File Binary File Format
 - [MS-XLS]: Excel Binary File Format (.xls) Structure Specification
 - [MS-XLSB]: Excel (.xlsb) Binary File Format
 - [MS-XLSX]: Excel (.xlsx) Extensions to the Office Open XML SpreadsheetML File Format
 - [MS-ODATA]: Open Data Protocol (OData)
 - [MS-OFFCRYPTO]: Office Document Cryptography Structure
 - [MS-OLEDS]: Object Linking and Embedding (OLE) Data Structures
 - [MS-OLEPS]: Object Linking and Embedding (OLE) Property Set Data Structures
 - [MS-OSHARED]: Office Common Data Types and Objects Structures
 - [MS-OVBA]: Office VBA File Format Structure
 - [MS-OE376]: Office Implementation Information for ECMA-376 Standards Support
 - [XLS]: Microsoft Office Excel 97-2007 Binary File Format Specification

Certain features are shared with the Office Open XML File Formats, covered in:

ISO/IEC 29500:2012(E) "Information technology — Document description and processing languages — Office Open XML File Formats"

## Badges

[![Build Status](https://travis-ci.org/SheetJS/js-xls.png?branch=master)](https://travis-ci.org/SheetJS/js-xls)

[![Coverage Status](https://coveralls.io/repos/SheetJS/js-xls/badge.png?branch=master)](https://coveralls.io/r/SheetJS/js-xls?branch=master)

[![githalytics.com alpha](https://cruel-carlota.pagodabox.com/4ee4284bf2c638cff8ed705c4438a686 "githalytics.com")](http://githalytics.com/SheetJS/js-xls)


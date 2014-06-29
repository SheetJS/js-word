# xls

Parser for Excel XLS (BIFF5/BIFF8) and 2003-2004 (XML) files.  Pure-JS cleanroom
implementation from the Microsoft Open Specifications and related documents.

## Installation

In [nodejs](https://www.npmjs.org/package/xlsjs):

    npm install xlsjs

In the browser:

    <script src="xls.js"></script>

In [bower](http://bower.io/search/?q=js-xls):

    bower install js-xls

CDNjs automatically pulls the latest version and makes all versions available at
<http://cdnjs.com/libraries/xls>

## Optional Modules

The nodejs version automatically requires modules for additional features.  Some
of these modules are rather large in size and are only needed in special
circumstances, so they do not ship with the core.  For browser use, they must
be included directly:

    <!-- international support from https://github.com/sheetjs/js-codepage -->
    <script src="dist/cpexcel.js"></script>

An appropriate version for each dependency is included in the dist/ directory.

The complete single-file version is generated at `dist/xls.full.min.js`

## ECMAScript 5 compatibility

Since xlsx.js uses ES5 functions like `Array#forEach`, older browsers require
[Polyfills](http://git.io/QVh77g).  This repo and the gh-pages branch include
[a shim](https://github.com/SheetJS/js-xls/blob/master/shim.js)

To use the shim, add the shim before the script tag that loads xlsx.js:

    <script type="text/javascript" src="/path/to/shim.js"></script>

## Parsing Workbooks

For parsing, the first step is to read the file.

- nodejs:

```
if(typeof require !== 'undefined') XLS = require('xlsjs');
var workbook = XLS.readFile('test.xls');
/* DO SOMETHING WITH workbook HERE */
```

- ajax:

```
/* set up XMLHttpRequest */
var url = "test_files/formula_stress_test_ajax.xls";
var oReq = new XMLHttpRequest();
oReq.open("GET", url, true);
oReq.responseType = "arraybuffer";

oReq.onload = function(e) {
  var arraybuffer = oReq.response;

  /* convert data to binary string */
  var data = new Uint8Array(arraybuffer);
  var arr = new Array();
  for(var i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);
  var bstr = arr.join("");

  /* Call XLS */
  var workbook = XLS.read(bstr, {type:"binary"});

  /* DO SOMETHING WITH workbook HERE */
}

oReq.send();
```

- html5 drag-and-drop using readAsBinaryString:

```
/* set up drag-and-drop event */
function handleDrop(e) {
  e.stopPropagation();
  e.preventDefault();
  var files = e.dataTransfer.files;
  var i,f;
  for (i = 0, f = files[i]; i != files.length; ++i) {
    var reader = new FileReader();
    var name = f.name;
    reader.onload = function(e) {
      var data = e.target.result;

      /* if binary string, read with type 'binary' */
      var wb = XLS.read(data, {type: 'binary'});

      /* DO SOMETHING WITH workbook HERE */
    };
    reader.readAsBinaryString(f);
  }
}
drop_dom_element.addEventListener('drop', handleDrop, false);
```

This example gets the value of cell A1 of the first worksheet:

```
var sheet_name_list = workbook.SheetNames;
var Sheet1A1 = workbook.Sheets[sheet_name_list[0]]['A1'].v;
```

Complete examples:

- <http://oss.sheetjs.com/js-xls/> HTML5 File API / Base64 Text / Web Workers

Note that older versions of IE does not support HTML5 File API, so the base64
mode is provided for testing.  On OSX you can get the base64 encoding with:

    $ <target_file.xls base64 | pbcopy

- <http://oss.sheetjs.com/js-xls/ajax.html> XMLHttpRequest

- <https://github.com/SheetJS/js-xls/blob/master/bin/xls.njs> nodejs

The nodejs version installs a binary `xls` which can read XLS and XML2003
files and output the contents in various formats.  The source is available at
`xls.njs` in the bin directory.

Some helper functions in `XLS.utils` generate different views of the sheets:

- `XLS.utils.sheet_to_csv` generates CSV
- `XLS.utils.sheet_to_json` generates an array of objects
- `XLS.utils.get_formulae` generates a list of formulae


## Interface

`XLS` is the exposed variable in the browser and the exported nodejs variable


`XLS.read(data, read_opts)` attempts to parse `data`.

`XLS.readFile(filename, read_opts)` attempts to read `filename` and parse.

## Cell Object Description

js-xls conforms to the Common Spreadsheet Format (CSF):

`.SheetNames` is an ordered list of the sheets in the workbook

`.Sheets[sheetname]` returns a data structure representing the sheet.  Each key
that does not start with `!` corresponds to a cell (using `A-1` notation).

`.Sheets[sheetname][address]` returns the specified cell:

- `.v` : the raw value of the cell
- `.w` : the formatted text of the cell (if applicable)
- `.t` : the type of the cell (constrained to the enumeration `ST_CellType` as
  documented in page 4215 of ISO/IEC 29500-1:2012(E) )
- `.f` : the formula of the cell (if applicable)
- `.z` : the number format string associated with the cell (if requested)

For dates, `.v` holds the raw date code from the sheet and `.w` holds the text

## Parsing Options

The exported `read` and `readFile` functions accept an options argument:

| Option Name | Default | Description |
| :---------- | ------: | :---------- |
| cellFormula | true    | Save formulae to the .f field ** |
| cellNF      | false   | Save number format string to the .z field |
| sheetRows   | 0       | If >0, read the first `sheetRows` rows ** |
| bookFiles   | false   | If true, add raw files to book object ** |
| bookProps   | false   | If true, only parse enough to get book metadata ** |
| bookSheets  | false   | If true, only parse enough to get the sheet names |
| password    | ""      | If defined and file is encrypted, use password ** |

- `cellFormula` only applies to constructing XLS formulae.  XLML R1C1 formulae
  are stored in plaintext, but XLS formulae are stored in a binary format.
- Even if `cellNF` is false, formatted text (.w) will be generated
- In some cases, sheets may be parsed even if `bookSheets` is false.
- `bookSheets` and `bookProps` combine to give both sets of information
- `bookFiles` adds a `cfb` object (XLS only)
- `sheetRows-1` rows will be generated when looking at the JSON object output
  (since the header row is counted as a row when parsing the data)
- Currently only XOR encryption is supported.  Unsupported error will be thrown
  for files employing other encryption methods.

## Tested Environments

Tests utilize the mocha testing framework.  Travis-CI and Sauce Labs links:

 - <https://travis-ci.org/SheetJS/js-xls> for XLS module in node
 - <https://travis-ci.org/SheetJS/SheetJS.github.io> for XLS* modules
 - <https://saucelabs.com/u/sheetjs> for XLS* modules using Sauce Labs

## Test Files

Test files are housed in [another repo](https://github.com/SheetJS/test_files).

Running `make init` will refresh the `test_files` submodule and get the files.

## Testing

`make test` will run the nodejs-based tests.  To run the in-browser tests, clone
[the oss.sheetjs.com repo](https://github.com/SheetJS/SheetJS.github.io) and
replace the xls.js file (then fire up the browser and go to `stress.html`):

```
$ cp xls.js ../SheetJS.github.io
$ cd ../SheetJS.github.io
$ simplehttpserver # or "python -mSimpleHTTPServer" or "serve"
$ open -a Chromium.app http://localhost:8000/stress.html
```

For a much smaller test, run `make test_misc`.

## Other Notes

`CFB` refers to the Microsoft Compound File Binary Format, a container format for XLS as well as DOC and other pre-OOXML data formats.

The mechanism is split into a CFB parser (which scans through the file and produces concrete data chunks) and a Workbook parser (which does excel-specific parsing).  XML files are not processed by the CFB parser.

## Contributing

Due to the precarious nature of the Open Specifications Promise, it is very
important to ensure code is cleanroom.  Consult CONTRIBUTING.md

## XLSX/XLSM/XLSB Support

XLSX/XLSM/XLSB is available in [js-xlsx](http://git.io/xlsx).

## License

Please consult the attached LICENSE file for details.  All rights not explicitly
granted by the Apache 2.0 license are reserved by the Original Author.

It is the opinion of the Original Author that this code conforms to the terms of
the Microsoft Open Specifications Promise, falling under the same terms as
OpenOffice (which is governed by the Apache License v2).  Given the vagaries of
the promise, the Original Author makes no legal claim that in fact end users are
protected from future actions.  It is highly recommended that, for commercial
uses, you consult a lawyer before proceeding.

## References

Certain features are shared with the Office Open XML File Formats, covered in:

ISO/IEC 29500:2012(E) "Information technology — Document description and processing languages — Office Open XML File Formats"

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
 - [MS-CTXLS]: Excel Custom Toolbar Binary File Format
 - [MS-XLDM]: Spreadsheet Data Model File Format
 - [XLS]: Microsoft Office Excel 97-2007 Binary File Format Specification

## Badges

[![Build Status](https://travis-ci.org/SheetJS/js-xls.svg?branch=master)](https://travis-ci.org/SheetJS/js-xls)

[![Coverage Status](http://img.shields.io/coveralls/SheetJS/js-xls/master.svg)](https://coveralls.io/r/SheetJS/js-xls?branch=master)

[![githalytics.com alpha](https://cruel-carlota.pagodabox.com/4ee4284bf2c638cff8ed705c4438a686 "githalytics.com")](http://githalytics.com/SheetJS/js-xls)

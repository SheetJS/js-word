# [SheetJS js-word](http://wordjs.com)

Parser and writer for various word processing doc formats. Pure-JS cleanroom
implementation from official specifications, related documents, and test files.
Emphasis on parsing and writing robustness, cross-format feature compatibility
with a unified JS representation, and maximal browser compatibility.


## Test Files

Test files should be placed in the `test_files` directory, in the appropriate
subdirectory for the filetype.  For example, DOCX files should be placed in
`test_files\docx\wordjs` and RTF files should be in `test_files\rtf\wordjs`.

Every test file should be accompanied by a plain text `.txt` representation
whose filename is the original filename appended with `.txt`.  For example, the
DOCX file `test_files\docx\wordjs\foo.docx` pairs with the plain text file `test_files\docx\wordjs\foo.docx.txt`

**Generating Baselines using Word for Windows**

0. Ensure you have PowerShell version 7.0 or greater
1. Run `Set-ExecutionPolicy RemoteSigned` OR `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` in Powershell (PS) Admin 7.0
2. Have the PS script in the root of the repo
3. Run `.\generate_txt.ps1 .\test_files\EXT_TYPE\FOLDER` (ex. `.\generate_txt.ps1 .\test_files\docx\apachepoi`)

On first run, if a test file does not have an accompanying `.txt` file, the
script will open Word and save the file as plaintext.  Word will rapidly open
and close during this process.

The script will not attempt to open Word or try to generate `.txt` files if they
already exist.  After a clean run, Word should not open on future runs.

The script will halt for documents that are broken in certain ways.  Word will
display a prompt, stalling the automated process.  Those documents can be
skipped by creating a `.skip` file as described below.


**Skipping Files**

The script will look for files with the `.skip` extension and skip processing
the base file.  For example, if `test_files\docx\wordjs\Hello.docx.skip` exists,
the script will not attempt to process `test_files\docx\wordjs\Hello.docx`

When the UI blocks (for example, on a VBA error with `ThisDocument`), the
corresponding `.skip` file should be created manually.  The script merely tests
if the file exists, so the content is immaterial and a single letter suffices.

**Generating `.skip` files**

The script will attempt to open password-protected documents using the password
"WordJS".  The script will not halt but it will not generate a text file.  With
a few adjustments, the script can generate `.skip` files for those cases

1. Uncomment [L27-29](https://github.com/SheetJS/js-word/blob/master/generate_txt.ps1#L27-L29) in the script
2. Comment [L26](https://github.com/SheetJS/js-word/blob/master/generate_txt.ps1#L26) in the script
3. Rerun the script
4. Undo Step 1 and 2


## License

Please consult the attached LICENSE file for details.  All rights not explicitly
granted by the Apache 2.0 License are reserved by the Original Author.


## References

<details>
  <summary><b>OSP-covered Specifications</b> (click to show)</summary>

 - `MS-CFB`: Compound File Binary File Format
 - `MS-DOC`: Word (.doc) Binary File Format
 - `RTF`: Rich Text Format

</details>

- ISO/IEC 29500:2012(E) "Information technology — Document description and processing languages — Office Open XML File Formats"
- Open Document Format for Office Applications Version 1.3 (25 December 2019)

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-word?pixel)](https://github.com/SheetJS/js-word)

# [SheetJS js-word](http://wordjs.com)

## Generating Test Files

If there is a test file that needs to be added, a plain text `.txt` generally needs to be created as well. 

**To run this script:**

1. Run `Set-ExecutionPolicy RemoteSigned` OR `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` in Powershell (PS) Admin 7.0
2. Have the PS script in the root of the repo
3. Run `.\generate_txt.ps1 .\test_files\EXT_TYPE\FOLDER` (ex. `.\generate_txt.ps1 .\test_files\docx\apachepoi`) 

*_Note: When generating more `.txt` files after the initial execution, Word should NOT be opening._

**If Word is opening, follow these steps:**

1. Uncomment [L27-29](https://github.com/SheetJS/js-word/blob/master/generate_txt.ps1#L27-L29) in the script 
2. Comment [L26](https://github.com/SheetJS/js-word/blob/master/generate_txt.ps1#L26) in the script
3. Rerun the script
4. Undo Step 1 and 2

_There should be `.skip` files for the password protected documents, and Word shouldn't open twice._

**Outputs that are expected from this script, are the following:**

- Word should be opening and closing (to grab and save the contents as .txt)
- There should be .txt files being generated in the same folder that you ran the script

**Edge cases to consider**:

- The script will fail for documents that are broken (Word will prompt a notification to handle this)
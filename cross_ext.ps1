$HelpCommand = "Run '.\generate_txt.ps1 -h' to see examples"

#Check if argument or option is provided
if ($Args.count -ne 1) {
    Write-Output "ERROR: Provide a test file directory or options. $HelpCommand"
    exit
}

# Check if argument is an option
if ($Args[0] -match '^-') {
    #Check if option is help
    if ($Args[0] -match '^-[h|H](elp)?$') {
        Write-Output "Usage:
        .\generate_txt[.ps1] <filePath>
        .\generate_txt[.ps1] -[h|H[elp]]
        Examples:
        filePath = .\test_files\docx\apachepoi
        "
    } else {
        Write-Output "ERROR: Provide a valid option. $HelpCommand"
        exit
    }
}

# Check if path exists
if (Test-Path $Args[0]) {
    $Directory = Resolve-Path $Args[0]
} else {
    Write-Output "ERROR: Provide a valid test file directory. $HelpCommand"
    exit
}

$Directory = Resolve-Path $Args[0]
$LineEnding = 1 # WdLineEndingType 1 = CROnly
$Encoding = 65001 # Codepage for UTF8
$parent = Split-Path (Split-Path $Directory -Parent) -Leaf
$subparent = Split-Path $Directory -Leaf

<# 
    Convert current extension to `.rtf`
#>
$CurrAbsPath = @(Get-ChildItem -path $Directory -Recurse -Exclude *.txt, *.skip)
:curr_main for($i=0; $i -lt $CurrAbsPath.length; $i++) {
    $CurrAbsPathI = Resolve-Path $CurrAbsPath[$i] | Select-Object -ExpandProperty Path
 
    # Check if file exists, if not continue
    try {
        $similarAbsPath = Join-Path -Path .\test_files\rtf -ChildPath $parent\$subparent
        $filename = Split-Path $CurrAbsPathI -Leaf
        $similarAbsFilePath = Resolve-Path (Join-Path -Path $similarAbsPath -ChildPath ($filename+".rtf")) -ErrorAction Stop

        if (Test-Path $similarAbsFilePath -PathType Leaf) {
            continue curr_main
        }
    } catch {
        continue curr_main
    }

    # if (Test-Path ($CurrAbsPathI + ".rtf") -PathType Leaf ) { continue curr_main }
 
    if (Test-Path ($CurrAbsPathI + ".skip")) { continue curr_main }
 
    # Save from current extension to `.rtf`
    $Word = New-Object -ComObject Word.Application
    try {
        $Doc = $Word.Documents.Open($CurrAbsPathI, $False, $True, $False, "WordJS", "WordJS")
        $Doc.SaveAs(($CurrAbsPathI+".rtf"), 6, $False, "", $False, "", $False, $False, $False, $False, $False, $Encoding, $False, $False, $LineEnding)
        $Doc.Close()
    } catch {
        Write-Output "Skipping (has pwd or cannot edit): $CurrAbsPathI"
    }
 
    Stop-Process -Name "winword"
}

<# 
    Save a `.txt` file of the `.rtf` file that 
    was just created. 

    Then mirror the folder structure into the `rtf/` folder
    and move the `.rtf` and `.txt` files into the subfolder.
#> 
$ExtAbsPath = @(Get-ChildItem -path $Directory -Recurse -Exclude *.txt, *.skip, *.docx)
$rtfPath = Join-Path -Path .\test_files\ -ChildPath rtf
:ext_main for($i=0; $i -lt $ExtAbsPath.length; $i++) {
    $ExtAbsPathI = Resolve-Path $ExtAbsPath[$i] | Select-Object -ExpandProperty Path

    # Only keep `.rtf` files that are under 10 mb
    if ((Get-Item $ExtAbsPathI).length -gt 10000kb) {
        Remove-Item $ExtAbsPathI
        continue ext_main
    }

    if (Test-Path ($ExtAbsPathI + ".txt") -PathType Leaf ) { continue ext_main }
 
    if (Test-Path ($ExtAbsPathI + ".skip")) { continue ext_main }
 
    # Mirror folder structure into `rtf/` 
    New-Item -Path $rtfPath -Name $parent -ItemType "directory" -Force
    $p_rtfPath = Join-Path -Path $rtfPath -ChildPath $parent
    New-Item -Path $p_rtfPath -Name $subparent -ItemType "directory" -Force
    $sp_rtfPath = Join-Path -Path $p_rtfPath -ChildPath $subparent

    # Save a `.txt` file of the `.rtf` 
    $Word = New-Object -ComObject Word.Application
    try {
        $Doc = $Word.Documents.Open($ExtAbsPathI, $False, $True, $False, "WordJS", "WordJS")
        $Doc.SaveAs(($ExtAbsPathI+".txt"), 7, $False, "", $False, "", $False, $False, $False, $False, $False, $Encoding, $False, $False, $LineEnding)
        $Doc.Close() 
        Move-Item -Path ($ExtAbsPathI+".txt") -Destination (Resolve-Path $sp_rtfPath) 
        Move-Item -Path ($ExtAbsPathI) -Destination (Resolve-Path $sp_rtfPath) 
    } catch {
        Write-Output "Skipping (has pwd): $ExtAbsPathI"
    }
 
    Stop-Process -Name "winword"
}
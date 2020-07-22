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

$LineEnding = 1 # WdLineEndingType 1 = CROnly
$Encoding = 65001 # Codepage for UTF8
# Read backlist
# $Blacklist = @()
# $parent = (get-item $Directory).parent.parent.parent
# $bl_file = Join-Path -Path $parent -ChildPath "blacklist.txt"
# foreach ($b in (Get-Content $bl_file)) {
#     $Blacklist += $b
# }

$AbsPath = @(Get-ChildItem -path $Directory -Recurse -Exclude *.txt, *.skip)
:main for($i=0; $i -lt $AbsPath.length; $i++) {
    $AbsPathI = Resolve-Path $AbsPath[$i] | Select-Object -ExpandProperty Path

    if (Test-Path ($AbsPathI + ".txt") -PathType Leaf ) { continue main }

    $filename = Split-Path $AbsPath[$i] -leaf
    # $join_fn = Join-Path -Path $Args[0] -ChildPath $filename
    <# if (($join_fn | Out-String) -contains $Blacklist) {
        Write-Output "Skipping $join_fn"
        continue main
    } #>

    if (Test-Path ($AbsPathI + ".skip")) { continue main }
    
    $Word = New-Object -ComObject Word.Application
    try {
        $Doc = $Word.Documents.Open($AbsPathI, $False, $True, $False, "WordJS", "WordJS")
        $Doc.SaveAs(($AbsPathI+".txt"), 7, $False, "", $False, "", $False, $False, $False, $False, $False, $Encoding, $False, $False, $LineEnding)
        $Doc.Close()
    } catch {
        $p = $Args[0]
        Get-Process | Out-File -FilePath ".\$p\$filename.skip"
        Write-Output "Skipping (has pwd or cannot edit): $AbsPathI"
        continue main
    }

    Stop-Process -Name "winword"
}
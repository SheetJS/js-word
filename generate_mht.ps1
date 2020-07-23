# Default path to ".\test_files\" if no argument is provided
if ($Args.Count -eq 0) {
    $pathString = ".\test_files\"
} else {
    $pathString = $Args[0]
}
$ext = ".mht"
$extno = 9
$Directory = Resolve-Path $pathString
$LineEnding = 1 # WdLineEndingType 1 = CROnly
$Encoding = 65001 # Codepage for UTF8
$testFileDepth = 2

# Exclude everything apart from .docx
$AbsPath = @(Get-ChildItem -path $Directory -Recurse -Exclude *.txt, *.skip, "*.mht", "*.rtf","*.odt","*.doc","*.html")
:mhtMain for ($i = 0; $i -lt $AbsPath.length; $i++) {

    $AbsPathI = Resolve-Path $AbsPath[$i] | Select-Object -ExpandProperty Path

    if (Test-Path ($AbsPathI + $ext) -PathType Leaf ) { continue mhtMain }

    if (Test-Path ($AbsPathI + ".skip")) { continue mhtMain }

    $Word = New-Object -ComObject Word.Application
    try {
        
        $Doc = $Word.Documents.Open($AbsPathI, $False, $True, $False, "WordJS", "WordJS")
        $Doc.SaveAs(($AbsPathI + $ext), $extno, $False, "", $False, "", $False, $False, $False, $False, $False, $Encoding, $False, $False, $LineEnding)
        $Doc.Close()
        
        # Variable examples:
        #   $AbsPathI       C:\Users\{Username}\Desktop\js-word\test_files\docx\apachepoi\52288.docx
        #   $list           ["C:", "Users", "{Username}","Desktop", "js-word","test_files","docx","apachepoi","52288.docx"]
        #   $index          5 (index of 'test_files)
        #   $secondHalf     apachepoi\52288.docx.mht
        #   $leaf           52288.docx
        #   $fileName       52288.docx.mht
        #   $mhtPath        .\test_files\mht\apachepoi\52288.docx.mht
        #   $currentPath    C:\Users\simpl\Desktop\pwsh_for_html\test_files\docx\apachepoi\52288.docx.mht
        
        $list = $AbsPathI -split '\\'
        $index = $list.IndexOf('test_files')
        $secondHalf = ($list | Select-Object -last ($list.Length - $index - $testFileDepth)) -join '\'
        $leaf = $secondHalf | Split-Path -Leaf
        $secondHalf = ($secondHalf | Split-Path -Parent) + "\" + $leaf + $ext
        $fileName = $leaf + $ext
        $mhtPath = Join-Path -Path "./test_files/mht" $secondHalf;
        $currentPath = (Split-Path $AbsPathI -Parent) + "\" + $fileName

        
        Write-Output $currentPath
        New-Item -ItemType File -Path $mhtPath -Force
        Move-Item -Path $currentPath -Destination $mhtPath -Force
    }
    catch {
        Write-Output "Skipping (has pwd): $AbsPathI"
    }

    Stop-Process -Name "winword"
}


$AbsPath = @(Get-ChildItem -path $Directory -Recurse -Exclude *.txt, *.skip, *.docx, "*.rtf","*.odt","*.doc","*.html")
:txtMain for ($i = 0; $i -lt $AbsPath.length; $i++) {
    $AbsPathI = Resolve-Path $AbsPath[$i] | Select-Object -ExpandProperty Path
    Write-Output $AbsPathI

    if (Test-Path ($AbsPathI + ".txt") -PathType Leaf ) { continue txtMain }

    if (Test-Path ($AbsPathI + ".skip")) { continue txtMain }

    $Word = New-Object -ComObject Word.Application
    try {
        $Doc = $Word.Documents.Open($AbsPathI, $False, $True, $False, "WordJS", "WordJS")
        $Doc.SaveAs(($AbsPathI + ".txt"), 7, $False, "", $False, "", $False, $False, $False, $False, $False, $Encoding, $False, $False, $LineEnding)
        $Doc.Close()
    }
    catch {
        Write-Output "Skipping (has pwd): $AbsPathI"
    }

    Stop-Process -Name "winword"
}
$ext = ".html"
$extno = 8
$Directory = Resolve-Path $Args[0]
$LineEnding = 1 # WdLineEndingType 1 = CROnly
$Encoding = 65001 # Codepage for UTF8
$testFileDepth = 2

$AbsPath = @(Get-ChildItem -path $Directory -Recurse -Exclude *.txt, *.skip, "*.html", "*.rtf","*.odt","*.doc")
:htmlMain for ($i = 0; $i -lt $AbsPath.length; $i++) {

    $AbsPathI = Resolve-Path $AbsPath[$i] | Select-Object -ExpandProperty Path

    if (Test-Path ($AbsPathI + $ext) -PathType Leaf ) { continue htmlMain }

    if (Test-Path ($AbsPathI + ".skip")) { continue htmlMain }

    $Word = New-Object -ComObject Word.Application
    try {
        $Doc = $Word.Documents.Open($AbsPathI, $False, $True, $False, "WordJS", "WordJS")
        $Doc.SaveAs(($AbsPathI + $ext), $extno, $False, "", $False, "", $False, $False, $False, $False, $False, $Encoding, $False, $False, $LineEnding)
        $Doc.Close()

        Remove-Item -Force -Recurse ($AbsPathI + "_files")

        $list = $AbsPathI -split '\\'
        $index = $list.IndexOf('test_files')
        $secondHalf = ($list | Select-Object -last ($list.Length - $index - $testFileDepth)) -join '\'
        $leaf = $secondHalf | Split-Path -Leaf
        $secondHalf = ($secondHalf | Split-Path -Parent) + "\" + $leaf + $ext
        $fileName = $leaf + $ext
        $htmlPath = Join-Path -Path "./test_files/html" $secondHalf;
        $current = (Split-Path $AbsPathI -Parent) + "\" + $fileName
        
        Write-Output $current
        New-Item -ItemType File -Path $htmlPath -Force
        Move-Item -Path $current -Destination $htmlPath -Force
    }
    catch {
        Write-Output "Skipping (has pwd)"
    }

    Stop-Process -Name "winword"
}


$AbsPath = @(Get-ChildItem -path $Directory -Recurse -Exclude *.txt, *.skip, *.docx, "*.rtf","*.odt","*.doc")
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
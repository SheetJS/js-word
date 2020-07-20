if ($Args.count -ne 1) {
    Write-Output "Provide a test file directory"
    exit
}
$Directory = Resolve-Path $Args[0]
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

    # $filename = Split-Path $AbsPath[$i] -leaf
    # $join_fn = Join-Path -Path $Args[0] -ChildPath $filename
    <# if (($join_fn | Out-String) -contains $Blacklist) {
        Write-Output "Skipping $join_fn"
        continue main
    } #>

    if (Test-Path ($AbsPathI + ".skip")) { continue main }
    # $p = $Args[0]
    # Get-Process | Out-File -FilePath ".\$p\$filename.skip"
    # continue main

    Start-Process winword.exe
    $Word = New-Object -ComObject Word.Application
    try {
        $Doc = $Word.Documents.Open($AbsPathI, $False, $True, $False, "WordJS", "WordJS")
        $Doc.SaveAs(($AbsPathI+".txt"), 7, $False, "", $False, "", $False, $False, $False, $False, $False, $Encoding, $False, $False, $LineEnding)
        $Doc.Close()
    } catch {
        Write-Output "Skipping (has pwd): $AbsPathI"
    }

    Stop-Process -Name "winword"
}
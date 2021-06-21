$ppt = New-Object -com powerpoint.application

$opt = [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF

Get-ChildItem '.\SMT1-ppt' -Filter '*.pptx' | ForEach-Object {
    $ifile = $_.FullName
    $pres = $ppt.Presentations.Open($ifile)
    $pathname = split-path $ifile
    $filename = Split-Path $ifile -Leaf
    $file = $filename.split(".")[0]
    $ofile = $pathname + "\..\SMT1-pdf\" + $file + ".pdf"
    $pres.SaveAs($ofile, $opt)
    $pres.close
}
$ppt.Quit

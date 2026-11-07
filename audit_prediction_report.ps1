Add-Type -AssemblyName System.IO.Compression.FileSystem
$doc = 'Aircraft_Tracker_Project_Report_Updated_With_Algorithm_Comparison.docx'
$zip = [System.IO.Compression.ZipFile]::OpenRead((Join-Path $PSScriptRoot $doc))
try {
  $entry = $zip.GetEntry('word\document.xml')
  $reader = [System.IO.StreamReader]::new($entry.Open())
  try { $xml = [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
} finally { $zip.Dispose() }
$ns=[System.Xml.XmlNamespaceManager]::new($xml.NameTable); $ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
$text=(($xml.SelectNodes('//w:t',$ns)|ForEach-Object{$_.InnerText}) -join ' ')
$checks=@(
 '3.6.1  Prediction Objective and Activation',
 '3.6.6  Justification for Selecting Sliding-Window Linear Regression',
 'Table 3.3: Comparison of suitable aircraft-trajectory prediction algorithms',
 'Table 3.4: Principal tools, technologies and their roles in the project',
 'Hybrid Kalman-ML model'
)
foreach($check in $checks){ Write-Output ("{0}: {1}" -f $check,$text.Contains($check)) }
$tables=$xml.SelectNodes('//w:body/w:tbl',$ns)
Write-Output ("Body tables: " + $tables.Count)
$comparison=$null
foreach($tbl in $tables){ if(((($tbl.SelectNodes('.//w:t',$ns)|ForEach-Object{$_.InnerText})-join ' ')) -like '*Sliding-window OLS linear regression*'){ $comparison=$tbl; break } }
if($comparison){
 $rows=$comparison.SelectNodes('./w:tr',$ns).Count
 $gridWidths=@($comparison.SelectNodes('./w:tblGrid/w:gridCol',$ns)|ForEach-Object{[int]$_.w})
 Write-Output ("Comparison table rows: $rows; columns: $($gridWidths.Count); grid sum: $(([int[]]$gridWidths|Measure-Object -Sum).Sum)")
}
Write-Output ("DOCX bytes: " + (Get-Item -LiteralPath $doc).Length)


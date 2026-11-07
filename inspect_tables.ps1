Add-Type -AssemblyName System.IO.Compression.FileSystem
$xml=[xml](Get-Content -LiteralPath 'report_updated_unpacked\word\document.xml' -Raw)
$ns=[System.Xml.XmlNamespaceManager]::new($xml.NameTable)
$ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
$i=0
foreach($tbl in $xml.SelectNodes('//w:body/w:tbl',$ns)){
  $i++
  $style=$tbl.SelectSingleNode('./w:tblPr/w:tblStyle',$ns)
  $rows=$tbl.SelectNodes('./w:tr',$ns).Count
  $cols=if($rows -gt 0){$tbl.SelectSingleNode('./w:tr[1]',$ns).SelectNodes('./w:tc',$ns).Count}else{0}
  $texts=@(); foreach($tc in $tbl.SelectNodes('./w:tr[1]/w:tc',$ns)){ $texts += (($tc.SelectNodes('.//w:t',$ns)|ForEach-Object{$_.InnerText}) -join '') }
  Write-Output ("{0}	style={1}	{2}x{3}	{4}" -f $i,$style.val,$rows,$cols,($texts -join ' | '))
}

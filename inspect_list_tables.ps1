$xml=[xml](Get-Content -LiteralPath 'report_updated_unpacked\word\document.xml' -Raw)
$ns=[System.Xml.XmlNamespaceManager]::new($xml.NameTable);$ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
$tbl=$xml.SelectNodes('//w:body/w:tbl',$ns)[1]
$i=0; foreach($tr in $tbl.SelectNodes('./w:tr',$ns)){ $i++; $cells=@(); foreach($tc in $tr.SelectNodes('./w:tc',$ns)){ $cells += (($tc.SelectNodes('.//w:t',$ns)|ForEach-Object{$_.InnerText})-join '') }; Write-Output ("$i	"+($cells-join ' | ')) }

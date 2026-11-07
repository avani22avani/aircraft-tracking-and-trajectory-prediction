$xml=[xml](Get-Content -LiteralPath 'report_updated_unpacked\word\document.xml' -Raw)
$ns=[System.Xml.XmlNamespaceManager]::new($xml.NameTable); $ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
foreach($target in @('The primary trajectory predictor','3.2.4  Compact Position Reporting','If fewer than four suitable samples')){
 $p=$xml.SelectNodes('//w:body/w:p',$ns)|Where-Object{(($_.SelectNodes('.//w:t',$ns)|ForEach-Object{$_.InnerText})-join '') -like "$target*"}|Select-Object -First 1
 Write-Output "--- $target ---"; if($p.pPr){Write-Output $p.pPr.OuterXml}else{Write-Output '(no pPr)'}
}

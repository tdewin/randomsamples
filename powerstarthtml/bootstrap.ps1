function dltextfile {
	param($uri,$fname)
	[Net.HttpWebRequest]$WebRequest = [Net.WebRequest]::Create($uri)
	[Net.HttpWebResponse]$WebResponse = $WebRequest.GetResponse()
	$Reader = New-Object IO.StreamReader($WebResponse.GetResponseStream())
	$Response = $Reader.ReadToEnd()
	$Reader.Close()
	
	$responsewin = $response -replace "(?<!\r)\n","`r`n"
	$responsewin | set-content -Encoding Utf8 -path $fname
}


if($Env:PSModulePath -match ('{0}[^\;]+' -f ($Env:USERPROFILE -replace '\\','\\'))) {
	$umodpath = ('{0}\PowerStartHTML' -f $Matches[0])
	$null = ni -Type Directory $umodpath -ErrorAction SilentlyContinue
	
	$uri = "https://raw.githubusercontent.com/tdewin/randomsamples/master/powerstarthtml/PowerStartHTML/PowerStartHTML.psm1"
	$fname = ('{0}\PowerStartHTML.psm1' -f $umodpath)
	dltextfile -uri $uri -fname $fname

	$uri = "https://raw.githubusercontent.com/tdewin/randomsamples/master/powerstarthtml/PowerStartHTML/PowerStartHTML.psd1"
	$fname = ('{0}\PowerStartHTML.psd1' -f $umodpath)
	dltextfile -uri $uri -fname $fname	

    write-host "All installed in $umodpath"
}
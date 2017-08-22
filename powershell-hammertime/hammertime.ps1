<#
	Keep trying connect to a port for (tcp timeout + $sleep)*$retry times
	
	Handy when you need to wait for a service to be online (or check if it is online) after for example a reboot
#>

function Connect-Port {
    [cmdletbinding()]
    param($hostname,$port,$sleep=10,$retry=10) 
    
    $hammertime = {
        param($hostname,$port) 
        try { 
            $r = ($s = [System.Net.Sockets.TcpClient]::new($hostname,$port)).Connected; 
            $s.Close(); 
            return $r 
        } 
        catch { return $false }
        return $false
    }
    $res = $false 
    while ($retry -gt 0 -and (-not ($res = $hammertime.Invoke($hostname,$port)))) {
        write-verbose "Retrying $retry"
        $retry -= 1
        start-sleep $sleep
    }  
    return $res
}
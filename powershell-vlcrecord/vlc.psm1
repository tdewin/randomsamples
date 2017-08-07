function Start-VLCRecording {
    [cmdletbinding()]
    param(
        [string]$vlcpath="",
        [string]$mousepath="",
        [string]$videopath=(join-path $env:USERPROFILE "Videos"),
        [string]$videoname=("vlc_record_{0}.mp4" -f (get-date).ToString("yyyyMMdd-hhmmss")),
        [int]$fps=30,
        [int]$videobitrate=0,
        [string]$transoverride=""
    )
    #--qt-start-minimized screen:// --screen-fps=30 :screen-mouse-image="E:\record\mouse.png" :sout=#transcode{vcodec=x264,vb=0,scale=0}:file{dst=c:\\recording\\record.mp4} :sout-keep

    if ($vlcpath -eq "") {
        $v64 = (join-path $env:ProgramFiles "VideoLAN\VLC\vlc.exe") 
        $v32 = (join-path ${env:ProgramFiles(x86)} "VideoLAN\VLC\vlc.exe")
        if((test-path $v64 )) {
            $vlcpath = $v64
        } elseif((test-path $v32)) {
            $vlcpath = $v32
        } else {
            try { 
                $vlcpath = "{0}\vlc.exe" -f (Get-ItemPropertyValue -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\VLC media player" -name "InstallLocationd"  -ErrorAction SilentlyContinue) }
            catch {
            }
        }
    }
    
    if($vlcpath -ne "") {
        $fullpath = join-path $videopath $videoname
        $transcodeln = ":sout=#transcode{vcodec=h264,vb=$videobitrate,scale=0}:file{dst=$fullpath}"

        if($transoverride -ne "") {
            $transcodeln = ":sout=#transcode{$transoverride}:file{dst=$fullpath}"
            write-verbose "Override $transcodeln"
        }

        $vlcargs = @(
            "--qt-start-minimized",
            "screen://",
            ("--screen-fps={0}" -f $fps),
            $transcodeln,
            ":sout-keep",
            "-I",
            "telnet",
            "--telnet-password=vlcrec",
            "--telnet-port=44224"
        )
        if($mousepath -ne "") {
            $vlcargs += ":screen-mouse-image=$mousepath"
        }
        write-verbose "Path $vlcpath"
        write-verbose "Args $vlcargs"
        Start-Process -FilePath $vlcpath -NoNewWindow -ArgumentList  $vlcargs
    } else {
        write-error "Could not find vlc.exe"
    }

}

function Stop-VLCRecording {
    param(
        [string]$server="localhost",
        [int]$port=44224
    )
    $socket = New-Object System.Net.Sockets.TcpClient($server, $port)
    if ($socket)
    {  
        $s = $socket.GetStream()
        $w = New-Object System.IO.StreamWriter($s,[System.Text.Encoding]::ASCII)
        $r = New-Object System.IO.StreamReader($s,[System.Text.Encoding]::ASCII)

        [char[]]$buf = New-Object char[] 1024

        for($da = $r.Peek();$da -gt 0;$da = $r.Peek()) {
            if($da -gt 1024) { $da = 1023 }
            $got = $r.read($buf,0,$da)
            #write-host ($buf[0..$got] -join "")
        }

        @("vlcrec","shutdown") | % {  $w.WriteLine($_);$w.Flush() }
    }
}
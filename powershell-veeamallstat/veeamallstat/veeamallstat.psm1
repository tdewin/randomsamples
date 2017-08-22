class VeeamAllStatJobSessionVM {
    $name
    $status;
    $starttime;
    $endtime;
    $size;
    $read;
    $transferred;
    $duration;
    $details;
    VeeamAllStatJobSessionVM() {}
    VeeamAllStatJobSessionVM($name,$status,$starttime,$endtime,$size,$read,$transferred,$duration,$details) {
        $this.name = $name
        $this.status = $status;
        $this.starttime = $starttime;
        $this.endtime = $endtime;
        $this.size = $size;
        $this.read = $read;
        $this.transferred = $transferred;
        $this.duration  = $duration;
        $this.details = $details
    }
}

class VeeamAllStatJobSession {
    $name = "<no name>" 
    $type = "<no type>";
    $description = "";
    $status = "<not run>";
    [System.DateTime]$creationTime='1970-01-01 00:00:00';
    [System.DateTime]$endTime='1970-01-01 00:00:00';
    [System.TimeSpan]$duration
    $processedObjects=0;
    $totalObjects=0
    $totalSize=0
    $backupSize=0
    $dataRead=0
    $transferredSize=0
    $dedupe=0
    $compression=0
    $details=""
    $vmstotal=0
    $vmssuccess=0
    $vmswarning=0
    $vmserror=0
    $allerrors=@{}
    [bool]$hasRan=$true
    [VeeamAllStatJobSessionVM[]]$vmsessions= @()
    VeeamAllStatJobSession() { }
    VeeamAllStatJobSession($name,$type,$description) {
        $this.name = $name
        $this.type = $type;
        $this.description =$description;
    }
    VeeamAllStatJobSession($name,$type,$description,$status,$creationTime,$endtime,$processedObjects,$totalObjects,$totalSize,$backupSize,$dataRead,$transferredSize,$dedupe,$compression,$details) {
        $this.name = $name
        $this.type = $type;
        $this.description =$description;
        $this.status = $status;
        $this.creationTime=$creationTime;
        $this.endTime=$endTime;
        $this.processedObjects=$processedObjects;
        $this.totalObjects=$totalObjects
        $this.totalSize=$totalSize
        $this.backupSize=$backupSize
        $this.dataRead=$dataRead
        $this.transferredSize=$transferredSize
        $this.dedupe=$dedupe
        $this.compression=$compression
        $this.details=$details
        $this.duration=$endTime-$creationTime
    }
}

class VeeamAllStatJobMain {
    $versionVeeam
    $server
    $serverString
    [VeeamAllStatJobSession[]]$jobsessions
    VeeamAllStatJobMain () {
        $this.jobsessions = @()
    }
}



function Get-VeeamAllStatJobSessionVMs {
    param(
        $session,
        [VeeamAllStatJobSession]$statjobsession
    )
    
    $tasks = $session.GetTaskSessions()

    foreach($task in $tasks) {
        $s = $task.status
        $vm = [VeeamAllStatJobSessionVM]::new(
            $task.Name,
            $s,
            $task.Progress.StartTime,
            $task.Progress.StopTime,
            $task.Progress.ProcessedSize,
            $task.Progress.ReadSize,
            $task.Progress.TransferedSize,
            $task.Progress.Duration,
            $task.GetDetails())

        if ($s -ieq "success") {
            $statjobsession.vmssuccess += 1 
        } elseif ($s -ieq "warning" -or $s -ieq "pending" -or $s -ieq "none") {
            $statjobsession.vmswarning +=1
        } else {
            $statjobsession.vmserror += 1
        }
        if ($vm.details -ne "") {
            $statjobsession.allerrors[$task.Name]=$vm.details
        }
        $statjobsession.vmsessions += $vm
        $statjobsession.vmstotal+=1
    }
}

function Get-VeeamAllStatJobSession {
    param(
        $job,
        $session
    )
    $statjob = [VeeamAllStatJobSession]::new(
        $job.Name,
        $session.JobType,
        $job.Description,
        $session.Result,
        $session.CreationTime,
        $session.EndTime,
        $session.Progress.ProcessedObjects,
        $session.Progress.TotalObjects,
        $session.Progress.TotalSize,
        $session.BackupStats.BackupSize,
        $session.Progress.ReadSize,
        $Session.Progress.TransferedSize,
        $session.BackupStats.GetDedupeX(),
        $session.BackupStats.GetCompressX(),
        $session.GetDetails()
    )
    Get-VeeamAllStatJobSessionVMs -session $session -statjobsession $statjob

    if ($session.Result -eq "None" -and $session.JobType -eq "BackupSync") {
        if($session.State -eq "Idle" -and $statjob.vmserror -eq 0 -and $statjob.vmswarning -eq 0 -and $statjob.allerrors.count -eq 0 -and $statjob.details -eq ""  -and $session.EndTime -gt $session.CreationTime ) {
            if ($session.Progress.Percents -eq 100) {
                $statjob.Status="Success"
            } 
        } 
    }

    return $statjob
}
function Get-VeeamAllStatJobSessions {
    param(
        [VeeamAllStatJobMain]$JobMain
    ) 

    $allsessions = Get-VBRBackupSession
    $allorderdedsess = $allsessions | Sort-Object -Property CreationTimeUTC -Descending  
    $jobs = get-vbrjob

    foreach ($Job in $Jobs) {
        $lastsession = $allorderdedsess | ? { $_.jobid -eq $Job.id } | select -First 1
        if ($lastsession -ne $null) {
           $JobMain.jobsessions += Get-VeeamAllStatJobSession -job $job -session $lastsession
        } else {
           $s = [VeeamAllStatJobSession]::new($job.Name,$job.type,$job.description)
           $s.hasRan = $false
           $JobMain.jobsessions += $s
        }
    }     

}

function Get-VeeamAllStatServerVersion {
    param(
        [VeeamAllStatJobMain]$JobMain
    )
    $versionstring = "Unknown Version"

    $pssversion = (Get-PSSnapin VeeamPSSnapin -ErrorAction SilentlyContinue)
    if ($pssversion -ne $null) {
        $versionstring = ("{0}.{1}" -f $pssversion.Version.Major,$pssversion.Version.Minor)
    }

    $corePath = Get-ItemProperty -Path "HKLM:\Software\Veeam\Veeam Backup and Replication\" -Name "CorePath" -ErrorAction SilentlyContinue
    if ($corePath -ne $null) {
        $depDLLPath = Join-Path -Path $corePath.CorePath -ChildPath "Packages\VeeamDeploymentDll.dll" -Resolve -ErrorAction SilentlyContinue
        if ($depDLLPath -ne $null -and (Test-Path -Path $depDLLPath)) {
            $file = Get-Item -Path $depDLLPath -ErrorAction SilentlyContinue
            if ($file -ne $null) {
                $versionstring = $file.VersionInfo.ProductVersion
            }
        }
    }

    $servsession = Get-VBRServerSession
    $JobMain.versionVeeam = $versionstring
    $JobMain.server = $servsession.server
    $JobMain.serverString = ("Server {0} : Veeam Backup & Replication {1}" -f $servsession.server,$versionstring)
}
function Get-VeeamAllStat {
    $report = [VeeamAllStatJobMain]::new()

    Get-VeeamAllStatServerVersion -JobMain $report
    Get-VeeamAllStatJobSessions -JobMain $report
    return $report
}

function Connect-TCPPort {
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



function Connect-VeeamAllStat {
    [cmdletbinding()]
    param(
        $allstatdir = $(join-path $env:programdata "veeamallstat"),
        $allstatfile = $(join-path $allstatdir "settings.xml"),
        $disconnect = $true
    )
    
    $null = New-Item -Type Directory $allstatdir -ErrorAction SilentlyContinue
    

    class VASSettings { $server;[int]$port;$username;$filename;VASSettings($s,$p,$u,$f) {$this.server = $s;$this.port=$p;$this.username=$u;$this.filename=$f} }
    $settings = [VASSettings]::new("127.0.0.1",9392,"administrator","myreport_%d.html")
    $password = ""
    $readfromfile = $false

    if(test-path -Type Leaf $allstatfile) {
        $in = Get-Content $allstatfile | ConvertFrom-Json
        $settings.server = $in.server
        $settings.port = $in.port
        $settings.username = $in.username
        $settings.filename = $in.filename
        $readfromfile = $true
    }

    [xml]$loginwpf = @"
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="Connect To" Height="375" Width="525">
    <Grid Height="324" VerticalAlignment="Top">
        <TextBox x:Name="txtsrv" HorizontalAlignment="Left" Height="23" Margin="105,161,0,0" TextWrapping="Wrap" Text="{0}:{1}" VerticalAlignment="Top" Width="388"/>
        <TextBox x:Name="txtuser" HorizontalAlignment="Left" Height="23" Margin="105,189,0,0" TextWrapping="Wrap" Text="{2}" VerticalAlignment="Top" Width="388"/>
        <PasswordBox x:Name="passwordBox" Height="23" HorizontalAlignment="Left" Margin="105,217,0,0"  VerticalAlignment="Top" Width="388"/>
        <TextBox x:Name="txtfile" HorizontalAlignment="Left" Height="23" Margin="105,245,0,0" TextWrapping="Wrap" Text="{3}" VerticalAlignment="Top" Width="388"/>
        <Label x:Name="labels" Content="Server" HorizontalAlignment="Left" Margin="33,161,0,0" VerticalAlignment="Top" Width="119" Height="23"/>
        <Label x:Name="labelu" Content="Username" HorizontalAlignment="Left" Margin="33,189,0,0" VerticalAlignment="Top" Width="119" Height="23"/>
        <Label x:Name="labelp" Content="Password" HorizontalAlignment="Left" Margin="33,217,0,0" VerticalAlignment="Top" Width="119" Height="23"/>
        <Label x:Name="header" Content="Veeam All Stat Report" HorizontalAlignment="Left" Height="102" Margin="33,26,0,0" VerticalAlignment="Top" Width="460" FontSize="40"/>
        <Button x:Name="btnLogin" Content="Create Report" HorizontalAlignment="Left" Margin="397,282,0,0" VerticalAlignment="Top" Width="96" IsDefault="True" Height="25"/>
        <Label x:Name="labelr" Content="Report" HorizontalAlignment="Left" Margin="33,245,0,0" VerticalAlignment="Top" Width="119" Height="35"/>
        <CheckBox x:Name="execute" Content="CheckBox" HorizontalAlignment="Left" Margin="33,306,0,0" VerticalAlignment="Top" Height="18" Width="131" Visibility="Hidden"/>
    </Grid>
</Window>
"@ -f $settings.server,$settings.port,$settings.username,$settings.filename
    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
    $reader=[System.Xml.XmlNodeReader]::new($loginwpf)
    $form = $null
    try{
        $form=[Windows.Markup.XamlReader]::Load( $reader )
    } 
    catch{Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."}

    $execute = $false
    if ($form -ne $null) {
        $btn = $form.FindName("btnLogin")
        $null = $btn.Add_Click({$form.FindName("execute").IsChecked = $true;$form.Close()})
        
        if($readfromfile) {
            $null = [System.Windows.Input.FocusManager]::SetFocusedElement($form,$form.FindName("passwordBox"))
        } else {
            $null = [System.Windows.Input.FocusManager]::SetFocusedElement($form,$form.FindName("txtserver"))
        }
        
        
        $null = $form.showDialog()
        $execute = $form.FindName("execute").IsChecked

        $server = $form.FindName("txtsrv").Text
        if($server -match "^([^:]+):([0-9]+)$") {
            $settings.server = $Matches[1]
            $settings.port = $Matches[2]
        } else {
            $settings.server = $server
        }
        $settings.username = $form.FindName("txtuser").Text
        $settings.filename =  $form.FindName("txtfile").Text
        $password = $form.FindName("passwordBox").Password

        
    }

    $settings | convertto-json | set-content $allstatfile
    #write-host $settings.server $settings.port $settings.username


    if( $execute -and $(Connect-TCPPort -hostname $settings.server -port $settings.port -verbose -sleep 1)) {
        add-pssnapin veeampssnapin

        $secstr = ConvertTo-SecureString -AsPlainText -Force $password
        $credentials = [System.Management.Automation.PSCredential]::new($settings.username,$secstr)

        Connect-VBRServer -server $settings.server -Credential  $credentials
        $fname = $settings.filename -replace "[%]d",(get-date).toString("yyyyMMdd-hhmm")
        write-verbose "Output to $fname"
        New-VeeamAllStatReport -path $fname
        if($disconnect) {
            Disconnect-VBRServer
        }
        & "explorer" $fname
    } else {
        write-verbose "Could not connect or cancelled"
    }

}

function Get-HumanDataSize {
 param([double]$numc)
 $num = $numc+0
 $trailing= "","K","M","G","T","P","E"
 $i=0
 while($num -gt 1024 -and $i -lt 6) {
  $num= $num/1024
  $i++
 }
 return ("{0:f1} {1}B" -f $num,$trailing[$i])
}

function Get-HumanDate {
    param([DateTime]$t)
    return $t.toString("yyyy-MM-dd HH:mm:ss")
}
function Get-HumanDuration {
    param([System.TimeSpan]$d)
    return ("{0:D2}:{1:D2}:{2:D2}" -f ($d.Hours+($d.Days*24)),$d.Minutes,$d.Seconds)
}
function New-VeeamAllStatReport {
    param(
        $path="report.html",
        $rpodate=(get-date -Hour 20 -Minute 0 -Second 0 -Millisecond 0).AddDays(-1)
    )

    $stats = Get-VeeamAllStat

    import-module PowerStartHTML
    $creationDate = (get-date).ToString("yyyy-MM-dd (ddd) HH:mm:ss")
   
    $ps = New-PowerStartHTML -title "My Job Report"
    $ps.Main().Add("div","jumbotron jumbotron-fluid").Add("div","container").N()
    $ps.Append("h1",$null,("Server {0}" -f $stats.server)).Append("hr","my-4").Append("p",$null,$creationDate).N()

    $ps.Main().Append("h2",$null,"Job Overview").N()
    $ps.cssStyles['.bgtitle'] = "background-color:#666666"
    $ps.Main().Add('table','table').Add("tr","bgtitle text-white").Append("th",$null,"Job").Append("th",$null,"Type").Append("th",$null,"Last Run").Append("th",$null,"Result").Up().N()

    $ps.cssStyles['.bgsubsection'] = "background-color:#eee;"
    $jss = $stats.jobsessions | sort-object -property name
    $sessionc = 0
    foreach ($s in $jss) {
        $statusclass = $runclass = "alert-danger"
        $lastrun = $s.creationTime

        if($lastrun -gt $rpodate) { $runclass = "alert-success"}
        if($s.status -eq "Success") { $statusclass = "alert-success"}
        
        if ($lastrun -lt (get-date -year 1980)) { $lastrun = "" }
        $ps.Add("tr").N()
        $ps.AddAttrs($ps.newEl,@{"data-toggle"="collapse";"data-target"="#subinfo-{0}" -f $sessionc})
        $ps.Append("td",$null,$s.Name).Append("td",$null,$s.type).Append("td",$runclass,$lastrun).Append("td",$statusclass,$s.status).Up().N()

        $ps.Add("tr","collapse").N()
        $ps.AddAttrs($ps.newEl,@{id="subinfo-{0}" -f $sessionc})
        if ($s.hasRan) {
            $ps.Add("td","bgsubsection").N()
           
            $ps.Add("table","table bgcolorsub").N()

            $ps.Add("tr").Append("td",$null,"Start Time").Append("td",$null,(Get-HumanDate $s.creationTime)).N()
                $ps.Append("td",$null,"VM Success").Append("td",$null,($s.vmssuccess)).Up().N()
            $ps.Add("tr").Append("td",$null,"End Time").Append("td",$null,(Get-HumanDate $s.endTime)).N()
                $ps.Append("td",$null,"VM Warning").Append("td",$null,($s.vmswarning)).Up().N()
            $ps.Add("tr").Append("td",$null,"Duration").Append("td",$null,(Get-HumanDuration $s.duration)).N()
                $ps.Append("td",$null,"VM Error").Append("td",$null,($s.vmserror)).Up().N()   

            if($s.details -ne "") {
                $ps.Add("tr").Append("td","alert-danger","Errors : {0}" -f $s.details).AddAttrs($ps.newEl,@{"colspan"=4})
                $ps.Up().N()
            }


            $ps.Up().Add("table","table bgcolorsub").N()
            $ps.Add("tr").Append("td",$null,"Total Size").Append("td",$null,(Get-HumanDataSize $s.totalSize)).N()
                $ps.Append("td",$null,"Total Objects").Append("td",$null,$s.totalObjects).Up().N()
            $ps.Add("tr").Append("td",$null,"Read Size").Append("td",$null,(Get-HumanDataSize $s.dataRead)).N()
                $ps.Append("td",$null,"Processed Objects").Append("td",$null,$s.processedObjects).Up().N()
            $ps.Add("tr").Append("td",$null,"Transferred Size").Append("td",$null,(Get-HumanDataSize $s.transferredSize)).N()
                $ps.Append("td",$null,"Compression").Append("td",$null,('{0}x' -f $s.Compression)).Up().N()
            $ps.Add("tr").Append("td",$null,"Target Size").Append("td",$null,(Get-HumanDataSize $s.backupSize)).N()
                $ps.Append("td",$null,"Dedupe").Append("td",$null,('{0}x' -f $s.Dedupe)).Up().N()


            $ps.Up().Add("table","table bgcolorsub").N()
            $ps.Add("tr").N()
            $headers = @("Name","Status","Start","Duration","Size","Read","Transfer")
            foreach ($h in $headers) {
                $ps.Append("th",$null,$h).N()
            }
            $ps.Up().N()
            foreach ($v in $s.vmsessions) {
                $statclass="alert-danger"
                if($v.status -eq "Success") { $statclass = "alert-success"} 
                
                $ps.Add("tr").N()
                $ps.Append("td",$null,$v.name).N()
                $ps.Append("td",$statclass,$v.status).N()
                $ps.Append("td",$null,(Get-HumanDate $v.startTime)).N()
                $ps.Append("td",$null,(Get-HumanDuration $v.duration)).N()
                $ps.Append("td",$null,(Get-HumanDataSize $v.size)).N()
                $ps.Append("td",$null,(Get-HumanDataSize $v.read)).N()
                $ps.Append("td",$null,(Get-HumanDataSize $v.transferred)).N()

                $ps.Up().N()

                if($v.details -ne "") {
                    $ps.Add("tr").Append("td","alert-danger","Errors : {0}" -f $v.details).AddAttrs($ps.newEl,@{"colspan"=$headers.count})
                    $ps.Up().N()
                }
            }
            $ps.Up().N()
        } else {
            $ps.Add("td","bgsubsection","Could not find a session, please make sure the job has already ran and/or is scheduled").N()
        }
        $ps.AddAttrs($ps.lastEl,@{"colspan"=4})
        $ps.Up().Up().N()
        $sessionc++
    }

    $ps.Main().Add("div",$null,"Veeam Version {0} " -f $stats.versionVeeam).N()
    $ps.onLoadJS.Add('window.showDetails = function(id) { alert(id) }') 
    $ps.Save($path)
}

Export-ModuleMember -Function Get-HumanDataSize
Export-ModuleMember -Function Get-HumanDate
Export-ModuleMember -Function Get-HumanDuration
Export-ModuleMember -Function Connect-VeeamAllStat
Export-ModuleMember -Function Get-VeeamAllStat
Export-ModuleMember -Function Get-VeeamAllStatJobSession
Export-ModuleMember -Function Get-VeeamAllStatServerVersion
Export-ModuleMember -Function New-VeeamAllStatReport

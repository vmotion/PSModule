#Dependancy on Con-EMU Terminal being installed. Remove -new_console switch (in putty command) to remove this dependancy
     


function Get-sInterfaceCust{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Hostname,
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        $custid
                   
    )
    if(!$creds){$global:creds = Get-Credential -UserName $env:USERNAME -Message 'Enter Password'}
    $option = [System.StringSplitOptions]::RemoveEmptyEntries
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($creds.Password)
    $pw = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    $computer = "$($hostname).peak10.net"
    New-SshSession -ComputerName $computer -Username $env:USERNAME -Password $pw
    $sshresult = Invoke-SshCommand -ComputerName $computer -Command "show interface desc | include $($custid)"
    $results = $sshresult | foreach {
        $parts = $_.split(" ", $option)
        New-Object -Type PSObject -Property([ordered]@{
            Interface = $parts[0]
            Status = "$($parts[1])/$($parts[2])"
            Desc = $parts[3]
            Switch = $computer
        })
    }
    return $results    
}

function Get-sInterface{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Hostname,
        [Parameter(Mandatory=$false,Position=1)]
        [boolean]$out                   
    )
if(!$creds){$global:creds = Get-Credential -UserName $env:USERNAME -Message 'Enter Password'}
    $option = [System.StringSplitOptions]::RemoveEmptyEntries
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($creds.Password)
    $pw = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    $a = Foreach ($h in $hostname){
            $computer = "$($h).peak10.net" 
            New-SshSession -ComputerName $computer -Username $env:USERNAME -Password $pw 
            $sshresult = Invoke-SshCommand -ComputerName $computer -Command "show interface description" -Quiet
            $option = [System.StringSplitOptions]::RemoveEmptyEntries
            $tempfile = "$($env:USERPROFILE)\$($computer).txt"
            $replace1 = 'Interface                      Status         Protocol Description' , ''
            $replace2 = 'admin down','ADown'
            $sshresult | where {$_ -ne ""}|set-content -path $tempfile
            $sshresult = get-content -path $tempfile | where {$_ -ne ""}
            $sshresult = $sshresult -replace $replace1 
            $sshresult = $sshresult -replace $replace2  
            $sshresult | set-content -path $tempfile
            $text = Get-Content -path $tempfile
            $custObj = New-Object PSObject
            $custObj =  $text | foreach {
                $parts = $_.split(" ", $option)
                New-Object -Type PSObject -Property([ordered]@{
                    Switch = $h
                    Interface = $parts[0]
                    Status = $parts[1]
                    Protocol = $parts[2]
                    Desc = $parts[3]
                    VLAN = if(!$parts[3]){}else{$parts[3].split(":")[2]}
                })
            }
            $sResults += $custObj |Where-Object {$_.interface -ne $null}
    }
    $ErrorActionPreference = 'SilentlyContinue'
    $sResults | foreach-object {$_ | Add-Member -MemberType NoteProperty -Name Suffix -Value (($_.desc).split(":"))[3]}
    $ErrorActionPreference = 'Continue'
    $sResults | Add-Member -MemberType ScriptMethod -Name 'ShowRun' -Value {
        param([string]$parameter01 = $(throw "Must supply an interface for to grab running config for. Example: Gi1/1"))
        $p = "show run interface $($parameter01)"
        $a = Get-sCommand -Hostname $this.switch -command $p
        return $a
    }
    $sResults | Export-Csv "S:\Provisioning\Central Provisioning\Automation\Inventory\SwitchExports\$($hostname)_Int.csv" -Force
    write-host 'Switch Interface output is saved to: S:\Provisioning\Central Provisioning\Automation\Inventory\SwitchExports' -ForegroundColor DarkYellow
    if($out -eq $true){
        return $sResults | Out-GridView -PassThru
    }else{return $sResults}
}

function Get-sInterfaceAll{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Hostname,

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        $password                   
    )
    if(!$password) {
        $creds = (Get-Credential -UserName $env:USERNAME -message "Enter Password for $($hostname)")
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($creds.Password)
        $pw = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        }
        else{
                $pw = $password
            }
    $a = Foreach ($h in $hostname){
            $computer = "$($h).peak10.net" 
            New-SshSession -ComputerName $computer -Username $env:USERNAME -Password $pw 
            $sshresult = Invoke-SshCommand -ComputerName $computer -Command "show interface description" -Quiet
            $option = [System.StringSplitOptions]::RemoveEmptyEntries
            $tempfile = "$($env:USERPROFILE)\$($computer).txt"
            $replace1 = 'Interface                      Status         Protocol Description' , ''
            $replace2 = 'admin down','ADown'
            $sshresult | where {$_ -ne ""}|set-content -path $tempfile
            $sshresult = get-content -path $tempfile | where {$_ -ne ""}
            $sshresult = $sshresult -replace $replace1 
            $sshresult = $sshresult -replace $replace2  
            $sshresult | set-content -path $tempfile
            $text = Get-Content -path $tempfile
            $custObj = New-Object PSObject
            $custObj =  $text | foreach {
                $parts = $_.split(" ", $option)
                New-Object -Type PSObject -Property([ordered]@{
                    Switch = $h
                    Interface = $parts[0]
                    Status = $parts[1]
                    Protocol = $parts[2]
                    Desc = $parts[3]
                    VLAN = if(!$parts[3]){}else{$parts[3].split(":")[2]}
                })
            }
            $sResults += $custObj |Where-Object {$_.interface -ne $null}
    }
    $sResults | Export-Csv "S:\Provisioning\Central Provisioning\Automation\Inventory\SwitchExports\$($hostname)_Int.csv" -Force
    write-host 'Switch Interface output is saved to: S:\Provisioning\Central Provisioning\Automation\Inventory\SwitchExports' -ForegroundColor DarkYellow
    return $sResults 
}

function Get-sCommand{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Hostname,

        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$false,
                   Position=1)]
        $command                  
    )
if(!$creds){$global:creds = Get-Credential -UserName $env:USERNAME -Message 'Enter Password'}
    $option = [System.StringSplitOptions]::RemoveEmptyEntries
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($creds.Password)
    $pw = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    $a = Foreach ($h in $hostname){
            $computer = "$($h).peak10.net"
            New-SshSession -ComputerName $computer -Username $env:USERNAME -Password $pw 
            $sshresult = Invoke-SshCommand -ComputerName $computer -Command $command -Quiet
            #$wshell = New-Object -ComObject Wscript.Shell
            #$wshell.Popup("$($SSHRESULT)",0,"Done",0x1)
    }
    return $sshResult
}


function get-CiscoVersion{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Hostname,
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$false,
                   Position=2)]
        $password 
    )

        $result = Get-sCommand -Hostname $hostname -password $password -command 'show version'
        $result > temp.txt
        $output = get-content temp.txt
        $newoutput = $output | Where-Object {$_ -like 'cisco *' -and $_ -notlike 'Cisco IOS*'}
        $version = $newoutput.split(" ")[1] 
        $newoutput1 = "$hostname, $version"    
        $newoutput1 >> Switch.txt
        Remove-Item .\Switch.txt
        remove-item temp.txt
return $newoutput1 
}

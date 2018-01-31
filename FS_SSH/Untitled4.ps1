#Dependancy on Con-EMU Terminal being installed. Remove -new_console switch (in putty command) to remove this dependancy
function Start-Putty{
        [CmdletBinding()]
        [Alias()]
        param(
        [Parameter(Mandatory=$True,Position=1)] 
        [string]$Hostname, 
        [Parameter(Mandatory=$False,Position=2)] 
        [String]$user,
        [Parameter(Mandatory=$False,Position=2)]
        [string]$password 
        )

    $user = $env:USERNAME
    $pass = if(!$password)  {
        $creds = get-credential -Message 'Enter Password for switch' -UserName $user
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($creds.password)
        $pass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        putty.exe -ssh -load $hostname -l $user -pw $Pass 
    }
    else{
        putty.exe -ssh -load $hostname -l $user -pw $password    
    }
}

function Start-cePutty{
        [CmdletBinding()]
        [Alias()]
        param(
        [Parameter(Mandatory=$True,Position=1)] 
        [string]$Hostname, 
        [Parameter(Mandatory=$False,Position=2)] 
        [String]$user,
        [Parameter(Mandatory=$False,Position=2)]
        [string]$password 
        )

    $user = $env:USERNAME
    $pass = if(!$password) {
        $creds = get-credential -Message 'Enter Password for switch' -UserName $user
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($creds.password)
        $pass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        putty.exe -new_console -ssh -load $hostname -l $user -pw $Pass 
    }else{
        putty.exe -new_console -ssh -load $hostname -l $user -pw $password    
     }
}      


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

    $option = [System.StringSplitOptions]::RemoveEmptyEntries
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($creds.Password)
    $pw = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    New-SshSession -ComputerName $hostname -Username $user -Password $pw
    $sshresult = Invoke-SshCommand -ComputerName $hostname -Command "show interface desc | include $($custid)"
    $results = $sshresult | foreach {
        $parts = $_.split(" ", $option)
        New-Object -Type PSObject -Property([ordered]@{
            Interface = $parts[0]
            Status = "$($parts[1])/$($parts[2])"
            Desc = $parts[3]
            Switch = $hostname
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
            New-SshSession -ComputerName $h -Username $env:USERNAME -Password $pw 
            $sshresult = Invoke-SshCommand -ComputerName $h -Command "show interface desc" -Quiet
            $option = [System.StringSplitOptions]::RemoveEmptyEntries
            $tempfile = "$($env:USERPROFILE)\$($h).txt"
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
    Remove-Item -path $tempfile -Force
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
        $command,       

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=2)]
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
            New-SshSession -ComputerName $hostname -Username $env:USERNAME -Password $pw 
            $sshresult = Invoke-SshCommand -ComputerName $h -Command $command -Quiet
            #$wshell = New-Object -ComObject Wscript.Shell
            #$wshell.Popup("$($SSHRESULT)",0,"Done",0x1)
    }
    return $sshResult
}


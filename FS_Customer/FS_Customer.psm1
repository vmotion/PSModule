function Get-netMktCustomers{
    [CmdletBinding()]
    param ( 
        [Parameter(Mandatory=$True,Position=1)] 
        [string]$market
        )
    $UriPre = 'https://overlook.peak10.com/api/device/'
    $UriPost = '/interface?limit=1000&extended_fetch=1'

    if(!$creds){$global:creds = Get-Credential -UserName $env:USERNAME -Message 'Enter Network Credentials'}
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($creds.password)
    $pass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

    $searchvalue = ":"
    $mktswitch = $switches | Where-Object {$_.market -eq $market -and $_.name -notlike '*Core*'}
    $mktswitches = $mktswitch | Where-Object {$_.em7_id -ne $null}
    $return = foreach($sw in $mktswitches){
        $em7_id = $sw.em7_id
        $uri = "$($UriPre)$($em7_id)$($UriPost)"
        $Results = Invoke-RestMethod -Method Get -Uri $Uri -Credential $creds
        $ints = $results.result_set |  GM -MemberType NoteProperty | select name
        foreach($int in $ints) {
           $results.result_set.$($int.name) | select @{N="Hostname";E={@($sw.Name)}}, name, alias, ifAdminstatus, ifoperstatus #| where-object {$_.alias -like "*$($searchvalue)*"} 
        }
    } 
    $regex = '[a-zA-z]+\d\d\d'
    $values = $return | foreach {$_.alias.split(":")[1]}
    $unique = $values | select -Unique | Select-String -AllMatches -Pattern $regex
    $results = $unique | Where-Object {$_.ToString().length -lt 9} | sort-object $_ | select -Unique
    return $results
}


function Get-CustomerMarkets{
    param([parameter(Mandatory=$true,ValueFromPipelineByPropertyName,ValueFromPipeline)][string]$custid)
    
        $a = import-csv 'S:\Provisioning\Central Provisioning\Automation\Inventory\SDInterfaces.csv'
        $b = import-csv 'S:\Provisioning\Central Provisioning\Automation\Inventory\Interfaces.csv'
        $c = import-csv 'S:\Provisioning\Central Provisioning\Automation\Inventory\vlans.csv'
        
        $a | where {$_.custid -eq $custid} | select -unique market
        $b | where {$_.custid -eq $custid} | select -unique market
        $c | where {$_.custid -eq $custid} | select -unique market

}

function Get-CustomerVLANs{
    [CmdletBinding()]
    param ( 
        [Parameter(Mandatory=$True,Position=1,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$CustID
        )
        $vlans = import-csv 'S:\Provisioning\Central Provisioning\Automation\Inventory\VLANs.CSV'
        $age = gci 'S:\Provisioning\Central Provisioning\Automation\Inventory\VLANs.CSV' | select LastWriteTime
        write-host "-----------------------------------------------------------------
        Using VLANs.CSV - Last Updated: $($age.LastWriteTime)
        Output saved to `$$($custid)_vlan
-----------------------------------------------------------------" -ForegroundColor Cyan
        
        $result = $vlans | where {$_.custid -eq $custid} 
        New-Variable -Name "$($custid)_VLAN" -Scope Global -Force -Value $result
        #$result | add-member -MemberType 

    return $result |select -unique CustID, Market, NetType, Suffix, VLAN, IPAddress| sort market, vlan| FT
}

function Get-CustomerInterfaces{
    [CmdletBinding()]
    param ( 
        [Parameter(Mandatory=$True,Position=1,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$CustID
        )
        $markets = Get-CustomerMarkets -custid $custid
        $interfaces = get-EM7Device -market $markets -SwitchType Access, DIST, SDDist, SD -PortType Interface
        $result = $interfaces | where {$_.custid -eq $custid}
        New-Variable -Name "$($custid)_Int" -Scope Global -Force -Value $result

        return $result
}

function lookup-custVLAN{
    param([parameter(Mandatory=$true,ValueFromPipelineByPropertyName,ValueFromPipeline)][string]$custid)

    $result = $vlans | where {$_.custid -eq $custid -and $_.hostname -like "*DIST-01*"}
    return $result
}

function Remove-CustomerNetworks{
        [CmdletBinding()]
    param ( 
        [Parameter(Mandatory=$True,Position=1)] 
        [string]$custid, 
        [Parameter(Mandatory=$True,Position=2)]
        [string]$markets
    )
    $intvarstring = "$($Custid)_Ints"
    write-host $intvarstring
    $vlanvarstring = "$($Custid)_VLANs"
    write-host $vlanvarstring


    $ints = get-eInterface -market $markets -SearchValue $custid

    #New-Variable -Name $intvarstring -Value $ints

    $vlans = (Get-Variable -Name $intvarstring).value
    #$vlans = $vlans| select -Unique hostname, vlan 
    #$vlans | foreach-object {$_ | add-member -MemberType ScriptMethod -Name Interfaces -value {
    #    $hInts = ($ints | where {$_.hostname -eq $this.hostname -and $_.vlan -eq $this.vlan} | select name).name
    #    return $hints
    #    }
    #}

    $vlans | foreach-object {$_ | Add-Member -force -MemberType NoteProperty -name PortChannel -value {
        $hn = $_.hostname 
        $PortChan = $switches | where {$_.name -eq $hn} | select Po
        return $PortChan.portchannel
    }}

    $backupConfigs = foreach($v in $vlans) {
        foreach($i in $v.interfaces){
        if(!$portchannel){
            Remove-AccessPort -hostname $i.hostname -interface $i.name -VLAN $i.vlan -PortChannel $portchannel
        }else{
            Remove-AccessPort -hostname $i.hostname -interface $i -VLAN $i.vlan
        }
        } 
    }

    $backupconfigs | FL Switch, interface, currentconfig | out-file backup.txt 
    $backupConfigs | FL Switch, interface, Commands | out-file commands.txt
    invoke-item Backup.txt
    invoke-item commands.txt
    return $ints

}

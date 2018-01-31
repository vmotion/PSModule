function get-eInterface{
    [CmdletBinding()]
    param ( 
        [Parameter(Mandatory=$True,Position=1,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)] 
        $market, 
        [Parameter(Mandatory=$false,Position=2)]
        [string]$SearchValue,
        [Parameter(mandatory=$false,Position=3,ValueFromPipelineByPropertyName=$true)]
        $SwitchType
        )
    $list = foreach($m in $market){
        $m1 = $m.Substring(0,3)
        $switches | where {$_.parentmarket -eq $m1} | select -unique name, em7_id, type, Po
    } 
    if(!$SwitchType){$slist = $list| select -unique Name, EM7_ID, Type}else{$slist = $list | select -unique Name, EM7_ID, Type| where {$_.type -in $SwitchType} | sort Type, Name}
    if($switchtype -eq 'ALL'){$slist = $list | select -unique Name, EM7_ID, Type}

    #EM7 URI
    $UriPre = 'https://overlook.peak10.com/api/device/'
    $UriPost = '/interface?limit=1000&extended_fetch=1'
    #Credentials
    if(!$creds){$global:creds = Get-Credential -UserName $env:USERNAME -Message 'Enter Network Credentials'}
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($creds.password)
    $pass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    $ErrorActionPreference = "SilentlyContinue"
    $counter = 0    
    $totalcount = $slist.count
    $EM7Result = foreach($sw in $slist){
        $em7_id = $sw.em7_id
        $uri = "$($UriPre)$($em7_id)$($UriPost)"
        Write-Progress -Activity "Searching for [ $($SearchValue) ] in Market: $($market)" -CurrentOperation $sw.name -PercentComplete (($counter / $totalcount) * 100) -Status "#$($counter) of $($totalcount)" 
        $Results = Invoke-RestMethod -Method Get -Uri $Uri -Credential $creds
        $ints = $results.result_set |  GM -MemberType NoteProperty | select name
        foreach($int in $ints) {
           if(!$SearchValue)
           {
                $results.result_set.$($int.name) | select @{N="Market";E={@($sw.Market)}}, @{N="Hostname";E={@($sw.Name)}},@{N="Type";E={@($sw.Type)}}, name, alias, @{N="Sts";E={@(Switch($_.ifoperstatus){1{"Up"}2{ "Down" }})}},  @{N="AdminStatus";E={@(Switch($_.ifAdminStatus){1{"Up"}2{ "Down" }})}}
           }else
           {
                $results.result_set.$($int.name) | select @{N="Market";E={@($sw.Market)}}, @{N="Hostname";E={@($sw.Name)}},@{N="Type";E={@($sw.Type)}}, name, alias, @{N="Sts";E={@(Switch($_.ifoperstatus){1{"Up"}2{ "Down" }})}},  @{N="AdminStatus";E={@(Switch($_.ifAdminStatus){1{"Up"}2{ "Down" }})}}| where-object {$_.alias -like "*$($searchvalue)*"} 
           }
        }$counter ++
    }
    $EM7Result |ForEach-Object{ $_ | Add-Member -MemberType ScriptProperty -name Status -value {[string]"$($this.sts)/$($this.adminstatus)"}}
    $index = 1
    $return = $EM7Result |Sort-Object Type, hostname, interface | where {$_.hostname -ne $null} | ForEach-Object {
        New-Object psobject -Property $AliasProperties
        $alias = $_.alias.split(":")
        $vl = $alias[2]
        $custid = $alias[1]
        $vlanid = $vl -replace 'V','' 
        $hname = $_.hostname.split("-")
        $vlanlookup = $vlans | where {$_.vlan -eq $vlanid -and $_.hostname -eq $hname[0]}
        $AliasProperties = [ordered]@{
            ID = $Index 
            hostname = $_.hostname
            name = $_.name
            Prefix = $alias[0]
            CustID= $custid
            VLAN= $vl
            Suffix= $alias[3]
            Status = $_.status 
            market = $hname[0]
            IPaddress = $vlanlookup.ipaddress
        }
        $index++
    }


## Additional Members
    $return | Add-Member -MemberType ScriptMethod -Name ShowRun -Value {
            $result = Get-sCommand -Hostname $this.hostname  -command "show run interface $($this.name)"
            $this | add-member -membertype NoteProperty -name ShowRun -value $result
            return $result
    }
    $return | add-member -MemberType ScriptMethod -name Methods -value { $this | GM -MemberType ScriptMethod }
    $return | Add-member -force -MemberType ScriptMethod -Name GetIntIP -value {
        $dists = $this | select -unique hostname |where {$_.hostname -like "*-DIST-0*"}
        $global:intIP = foreach($d in $dists){
            $rset = Get-sCommand -Hostname $d.hostname -command "sh ip interface brief"
            $r1 = ($rset -split '[\r\n]') |? {$_}
            $distVLANs = $this | select hostname, vlan, name | where {$_.name -like "vl*" -and $_.hostname -like "*-DIST-0*"}
            $vl = $distvlans | select name,vlan, @{N="VLANID";E={@(($_.vlan).substring(1,($_.vlan).Length - 1) )}}
            $regex =  '\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b'
            $IP_Lookup = foreach($v in $vl){
                $z = $r1 | where {$_ -like "Vlan$($v.vlanid)*"}
                $split = $z.split(" ",[System.StringSplitOptions]::RemoveEmptyEntries)
                $vname = $split[0]
                $ip = $split[1]
                $props = @{
                    vName = $vname
                    IPAddress = $ip
                }
                $IP_Lookupet = New-Object psobject -Property $props
                $IP_Lookupet  
            }
            $IP_Lookup | select @{N="VLAN";E={@("V"+$_.vName.substring(4,(($_.vname).length - 4)))}}, IPaddress
        }
        foreach($i in $intIP){    
            ($this | where {$_.vlan -eq $i.vlan}).IPaddress = $i.IPaddress
        }
        return $intIP 
    }
    $return | add-member -force -MemberType ScriptMethod -Name GetInterface -Value {
        get-interface -hostname $this.hostname -name $this.name
    }
    $return | add-member -MemberType ScriptMethod -Name ChildSwitches -Value  {
        $vlanid = $($this.VLAN.substring(1,$this.vlan.Length - 1))
        $r1= Get-sCommand -Hostname $this.hostname -command "show vlan id $($vlanid)"
        $POValues = Select-String 'Po\d\d\d' -input $r1 -AllMatches | Foreach {$_.matches} | select @{N="PortChannel";E={@($_.value)}}
        $childSwitches = ($list | where {$_.po -in $POValues.portchannel}).Name 
        $return | a
    }

    $defaultProperties = @('ID', 'VLAN','Hostname','Name')
    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultProperties)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    $return | Add-Member MemberSet PSStandardMembers $PSStandardMembers
    $varName = "$($searchvalue)_Ints"
    new-variable -Name $varName -Value $return -Scope Global 
    write-host "Object saved to variable: $($varname)"
    #$return | convertto-html 
    return $return | FT *
}

function get-marketDevices{
    [CmdletBinding()]
    param ( 
        [Parameter(Mandatory,Position=1,ValueFromPipelineByPropertyName)] 
        $market,
        [Parameter(Mandatory=$false)]$DeviceType
        )
 
        $result = foreach($m in $market){
            write-host $m
            if(!$DeviceType){
            $switches | Where-Object {$_.market -eq $m} 
            }else{
                 $switches | Where-Object {$_.market -eq $m}| where {$_.type -in $DeviceType}
            }
        }
        return $result | Select-Object ParentMarket, Market, Name, EM7_ID, Type| Sort-Object Type

}

function get-sVLAN{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        [Parameter(Mandatory=$True,ValueFromPipelineByPropertyName=$true,Position=1)]
        $Hostname,
        [Parameter(Mandatory=$false,Position=2)]
        [string]$Range

    )

    Begin
    {
        $host.privatedata.ProgressForegroundColor = "yellow"
        $host.privatedata.ProgressBackgroundColor = "black";
        $counter = 0
        $ScriptStart1 = (Get-Date)
        $totalcount = $hostname.count
        if(!$Range)
        {
            $VLANtypes = import-csv "S:\Provisioning\Central Provisioning\Automation\Inventory\VLAN_Ranges.csv"
            $VLAN = $vlantypes | Out-GridView -PassThru
            $command = "show vlan id $($VLAN.range) | include :"
            Write-Debug $command
            if([regex]::Matches($vlan.range,'\b\d{3,4}[-]\d{3,4}').success -eq $true)
            {
                #write-host ($($VLAN.range.split("-")[0])..$($VLAN.range.split("-")[1]))
                $rng = ($($VLAN.range.split("-")[0])..$($VLAN.range.split("-")[1]))
                $RangeDisplay = $VLAN.range
            }
        }else{
            $command = "show vlan id $($Range) | include :"
            Write-Debug $command
            if([regex]::Matches($range, '\b\d{3,4}[-]\d{3,4}').success -eq $true){
                $rng = ($($range.split("-")[0])..$($range.split("-")[1]))
                $RangeDisplay = $range
            }
         }
    }
    Process
    {
        foreach($h in $hostname){
            Write-Verbose $h
            Write-Progress -Activity 'Collecting VLANs' -CurrentOperation $h -PercentComplete (($counter / $totalcount) * 100) -Status "#$($counter) Total: $($totalcount)"
            $return = Get-sCommand -Hostname $h -command $command    
            $r1 = ($return -split '[\r\n]') |? {$_} 
            $r2 = $r1 -replace ' {2,}', ","
            $r3 = $r2 -replace ' ', ","
            $details = foreach($i in $r3){
                $obj = new-object psobject
                $vlan = [int]($i.split(",")[0])
                $desc = $i.split(",")[1]
                $status = $i.split(",")[2]
                $custid = $desc.split(":")[1]
                $market= $h.split("-")[0]
                $networkType = $desc.split(":")[0]
                $suffix = $desc.split(":")[3]
                $obj | Add-member -MemberType NoteProperty -Name Market -value $market 
                $obj | Add-member -MemberType NoteProperty -Name VLAN -value $VLAN
                $obj | Add-member -MemberType NoteProperty -Name CustID -Value $custid
                $obj | Add-member -MemberType NoteProperty -Name Desc -value $desc
                $obj | Add-member -MemberType NoteProperty -Name Status -value $Status
                $obj | add-member -MemberType NoteProperty -Name NetType -value $networkType 
                $obj | add-member -MemberType NoteProperty -Name Suffix -value $suffix
                $obj | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                $defaultProperties = @('Market','VLAN','CustID','Desc')
                $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultProperties)
                $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
                $obj | add-member -MemberType ScriptMethod -Name GetInt -Value {
                    get-distinterfaces -SearchValue $this.vlan -distSwitch ((get-market -Market $this.market).dist1)
                } 
                $obj
            }
            $counter ++
            $details
        }
    }
    End
    {
        $ScriptEnd1 = (Get-Date)
        $RunTime1 = New-Timespan -Start $ScriptStart1 -End $ScriptEnd1
        write-Verbose (“Execution Time for Dist: {0}m:{1}s” -f [math]::abs($Runtime1.Minutes),[math]::abs($RunTime1.Seconds)) 
    }
}

function new-VLAN_Trunk{
    [CmdletBinding()]
    param ( 
        [Parameter(Mandatory=$True,Position=1)] 
        [string]$SwitchName, 
        [Parameter(Mandatory=$True,Position=2)]
        [string]$PortChannel,
        [parameter(Mandatory=$True,position=3)]
        [string]$vlan

        )

    $Po_template = "
# $($SwitchName)

interface Port-channel$portchannel
switchport trunk allowed vlan add $vlan"

return $Po_template
}

function get-DistInterfaces{
    [CmdletBinding()]
    param ( 
        [Parameter(Mandatory=$false,Position=1)]
        [string]$SearchValue,
        [parameter(Mandatory=$false,position=2)]
        [System.Object]$distSwitch
        ) 
    $UriPre = 'https://overlook.peak10.com/api/device/'
    $UriPost = '/interface?limit=1000&extended_fetch=1'
    
    if(!$creds){$global:creds = Get-Credential -UserName $env:USERNAME -Message 'Enter Network Credentials'}
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($creds.password)
    $pass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    
    if(!$distSwitch){$dists = $switches | where-object {$_.type -eq 'Dist' -and $_.em7_id -ne ''}}else{$dists = $switches | where-object {$_.name -eq $distSwitch}}
    
    if(!$searchvalue) {$searchvalue = read-host -Prompt 'Value to Search'} 
    $distInts = foreach($sw in $dists){
        $em7_id = $sw.em7_id
        $uri = "$($UriPre)$($em7_id)$($UriPost)"
        $Results = Invoke-RestMethod -Method Get -Uri $Uri -Credential $creds
        $ints = $results.result_set |  GM -MemberType NoteProperty | select name
        foreach($int in $ints) {
           $results.result_set.$($int.name) | select @{N="Hostname";E={@($sw.Name)}}, name, alias, @{N="Status";E={@(Switch($_.ifoperstatus){1{"Up"}2{ "Down" }})}},  @{N="AdminStatus";E={@(Switch($_.ifAdminStatus){1{"Up"}2{ "Down" }})}}, ifAdminstatus, ifoperstatus| where-object {$_.alias -like "*$($searchvalue)*"} 
        }
    }
#VLAN ID
    $distints | foreach {$_ | Add-Member -MemberType NoteProperty -name VLAN -value ([string]($_.alias.split(":")[2].substring(1,3)))}
#Status
    $distints | add-member -MemberType NoteProperty -name 'PortStatus' -value "$($_.status)/$($_.adminstatus)"
#Market
    $distints |  ForEach-Object {$_ | add-member -MemberType NoteProperty -name 'Market' -value ($($_.hostname).split("-"))[0].substring(0,3)}

#Show interface
    $distints | Add-Member -MemberType ScriptProperty -Name ShowInt -Value {
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($creds.password)
            $pass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
            $result = Get-sCommand -Hostname $this.hostname  -password $pass -command "show interface $($this.name)"
            $results = ($result -split '[\r\n]') |? {$_}
            return $results
    }

#Show run interface
    $distints | Add-Member -MemberType ScriptProperty -Name ShowRun -Value {
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($creds.password)
            $pass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
            $result = Get-sCommand -Hostname $this.hostname  -password $pass -command "show run interface $($this.name)"
            $results = ($result -split '[\r\n]') |? {$_}
            return $results
    }

#Port-Channel
    $distints | add-member -MemberType ScriptProperty -Name PortChannel -Value  {
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($creds.password)
            $pass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
            $result = Get-sCommand -Hostname $this.hostname -command "show vlan id $($this.VLAN)"
            $POValues = Select-String 'Po\d\d\d' -input $result -AllMatches | Foreach {$_.matches} | select @{N="PortChannel";E={@($_.value)}}, @{N="S1";E={@($_.value.length-2)}}
            $POValues | ForEach-Object {$_ | add-member -MemberType NoteProperty -name 'SwitchNum' -value $_.portchannel.substring($_.S1,2) }
            $POValues | add-member -MemberType NoteProperty -name 'Switch' -value ($switches |Where-Object {$_.type -eq 'ACCESS' -and $_.market -like "$($market)*" -and $_.name -like "*-$($switchnum)"})
            return $POValues 
    }
    return $distInts

#Default View
    $defaultProperties = @('Hostname','Name','Alias','VLAN')
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultProperties)
    $distints | Add-Member MemberSet PSStandardMembers $PSStandardMembers
}

function get-FreePorts{
    [CmdletBinding()]
    param ( 
        [Parameter(Mandatory=$True,Position=1)] 
        [market]$market, 
        [parameter(Mandatory=$false,position=3)]
        [System.Object]$credentials,
        [parameter(Mandatory=$True,position=2)]
        [ValidateSet("Access","Core","DIST","SD","Edge", "All")]
        [String]$SwitchType
        )

 
    $UriPre = 'https://overlook.peak10.com/api/device/'
    $UriPost = '/interface?limit=1000&extended_fetch=1'
    
    Enum Market
    {    
    ATL = 1
    ATL1 = 2
    ATL2 = 3
    ATL3 = 4
    CIN = 5
    CIN1 = 6
    CIN2 = 7
    CLT = 8
    CLT1 = 9
    CLT2 = 10
    CLT3 = 11
    CLT4 = 12
    FLL = 13
    FLL1 = 14
    FLL2 = 15
    JAX = 16
    JAX1 = 17
    JAX2 = 18
    LOU = 19
    LOU1 = 20
    LOU2 = 21
    LOU3 = 22
    LOU4 = 23
    NAS = 24
    NAS1 = 25
    NAS4 = 26
    NAS2 = 27
    NAS3 = 28
    NAS5 = 29
    RAL = 30
    RAL1 = 31
    RAL2 = 32
    RAL3 = 33
    RIC = 34
    RIC1 = 35
    RIC2 = 36
    TPA = 37
    TPA1 = 38
    TPA2 = 39
    TPA3 = 40
    LAB = 41
    }

    if(!$creds){$global:creds = Get-Credential -UserName $env:USERNAME -Message 'Enter Network Credentials'}
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($creds.password)
    $pass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

    $searchvalue = 'v998'
    if($switchtype -ne "All"){$switchT = $switches | where-object {$_.type -eq $SwitchType}}else{$switchT = $switches}
    $switch = $switchT | where-object {$_.market -eq $market -and $_.type -ne 'Core' -and $_.em7_id -ne ''}
    $return = foreach($sw in $switch){
        $em7_id = $sw.em7_id
        $uri = "$($UriPre)$($em7_id)$($UriPost)"
        $Results = Invoke-RestMethod -Method Get -Uri $Uri -Credential $creds
        $ints = $results.result_set |  GM -MemberType NoteProperty | select name
        foreach($int in $ints) {
           $results.result_set.$($int.name) | select @{N="Hostname";E={@($sw.Name)}}, name, alias, ifAdminstatus, ifoperstatus| where-object {$_.alias -like "*$($searchvalue)*"} 
        }
    }   

    $return | Add-member -MemberType ScriptMethod -Name SwitchDesc -value {
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($creds.password)
            $pass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
            $result = Get-sCommand -Hostname $this.hostname  -password $pass -command "show interface description"
            return $result
    }

    #Market
    $return |  ForEach-Object {$_ | add-member -MemberType NoteProperty -name 'Market' -value ($($_.hostname).split("-"))[0].substring(0,3)}
    #Default View
    $defaultProperties = @('Market','Hostname','Name')
    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultProperties)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    $return | Add-Member MemberSet PSStandardMembers $PSStandardMembers
    return $return 

} 

function Get-VLAN_PO{
    [CmdletBinding()]
    param ( 
        [Parameter(Mandatory=$false,Position=1)]
        [string]$market,
        [parameter(Mandatory=$false,position=2)]
        [System.Object]$vlan,
        [parameter(position=3)]
        [string]$network
        )
        if($network = 'SD'){$hostname = (get-market -market $market).sddist1}else{
        $hostname = (get-market -Market $market).dist1}
        write-host $hostname
        $marketsplit = $hostname.split("-")[0]
        $aswitches = $switches | where {$_.type -eq 'access' -and $_.name -like "*$($marketsplit)*"}
#Port-Channel
            $result = Get-sCommand -Hostname $hostname -command "show vlan id $($vlan)"
            $POValues = Select-String '(Po\d{1,4})' -input $result -AllMatches | Foreach {$_.matches} | select @{N="PortChannel";E={@($_.value)}}
            #$POValues | ForEach-Object {$_ | add-member -MemberType NoteProperty -name 'SwitchNum' -value $_.portchannel.substring($_.S1,2) }
            $s = $aswitches | where {$_.Po -in $POValues.portchannel} | select name, Po
            $s | foreach-object {$_ | Add-Member -MemberType NoteProperty -Name 'VLAN' -Value "$($vlan)"}
            return $s
}

function Get-sIntConfig { 
    [CmdletBinding()]
    param ( 
        [Parameter(Mandatory=$false,Position=1, ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$hostname,
        [parameter(Mandatory=$false,position=2,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [Alias("Name")]
        [string]$interface
        )
        $command = "show run int $($interface)"
        $result = Get-sCommand -Hostname $hostname -command $command
        return $result            

}

function New-AccessPort{
    [CmdletBinding()]
    param ( 
        [Parameter(Mandatory=$True,Position=1)] 
        [string]$SwitchName, 
        [Parameter(Mandatory=$True,Position=2)]
        [string]$AccessPort,
        [parameter(Mandatory=$True,position=3)]
        [string]$vlan,
        [parameter(Mandatory=$True,position=4)]
        [String]$custid,
        [parameter(Mandatory=$True,position=7)]
        [String]$ratelimit,
        [parameter(Mandatory=$false,position=9)]
        [String]$cname,
        [parameter(Mandatory=$true,position=8)]
        [ValidateSet("Int","SD","Access")]
        [string]$type,
        [parameter(Mandatory=$false,position=9)]
        [string]$suffix

        )

    $header = "!"+"#"*60
    $row2_1 = "#"*[int](30 - ($SwitchName.Length /2)-3)+" $($switchname) "
    $row2_2 = "#"*(60-$row2_1.Length)
    $Banner = "!$($row2_1)$($row2_2)"
    $INT = Get-Content $PSScriptRoot\CloudPort.txt| Out-String 
    $sd = Get-Content $PSScriptRoot\SDPort.txt| Out-String 
    $access = Get-Content $PSScriptRoot\accessport.txt| Out-String 

    $portconfig = switch ($type)
    {
        'INT' {$int}
        'SD' {$sd}
        'Access' {$access}
    }


$portconfig1 = Invoke-Expression "`"$portconfig`""

$config = "
$($header)
$($banner)

$($portconfig1)

$($banner)
$($header)"

$emailtemplate="
$($header)

 Customer: $($cname), $($CustID)

    VLAN:           V$($VLAN)
    Access Switch:  $($Switchname)
    Port:           $($Accessport)
    Speed/Duplex:   Auto/Auto

$($header)

"


    $Result = New-Object psobject
    $result | add-member -MemberType NoteProperty -name 'CustID' -value $custid
    $result | add-member -MemberType NoteProperty -name 'Email' -value $emailtemplate
    $result | add-member -MemberType NoteProperty -name 'VLAN' -value $vlan
    $result | add-member -MemberType NoteProperty -Name 'Access Switch' -Value $SwitchName
    $result | Add-member -MemberType NoteProperty -name 'Port' -value $AccessPort
    $result | Add-Member -MemberType NoteProperty -name 'Config' -value $config
    $result | Add-Member -MemberType ScriptMethod -name Outfile -value {
       $file = "H:\$($this.CustID)_V$($this.VLAN)_$($this.'Access Switch').txt"
       write-host $file
       $this | fl email, config, potemplate | out-string | out-file -FilePath $file
       invoke-item $file
    }
    $mktCode = ($SwitchName.split("-"))[0]
    $market = get-market -Market $mktCode
    $Po = "1$($SwitchName.split("-")[2])"
    $row2_1 = "#"*[int](30 - ((($market.dist1).Length + ($market.dist2).Length) /2)-3)+" $($market.dist1) &  $($market.dist2) "
    $row2_2 = "#"*(60-$row2_1.Length)
    $Banner = "!$($row2_1)$($row2_2)"

$portChannelTemplate = "
$($header)
$($Banner)


interface Port-channel$($Po)
switchport trunk allowed vlan add $($VLAN)


$($Banner)
$($header)"
    
    $result | add-member -MemberType NoteProperty -name 'PoTemplate' -value $portChannelTemplate 
        
    return $result
}

function Remove-AccessPort{
        [CmdletBinding()]
    param ( 
        [Parameter(Mandatory=$True,Position=1)] 
        [string]$hostname, 
        [Parameter(Mandatory=$True,Position=2)]
        [string]$interface,
        [Parameter(Mandatory=$True,Position=3)]
        [string]$vlan,
        [Parameter(Mandatory=$false,Position=4)]
        [string]$PortChannel
    )

    $currentConfig = Get-sIntConfig -hostname $hostname -interface $interface 
    
    if($interface -like "VL*"){
        $dtemp = 
@"
        No interface $($vlan.Substring(1, $vlan.Length - 1))
        exit
        
        interface $($PortChannel)
         switchport trunk allowed vlan remove $($vlan.Substring(1, $vlan.Length - 1))
        exit
        end
        
"@
    }
        else{
    $dtemp = @"
#$($hostname)
	default interface $($interface)
	interface $($interface)
	    description SD::V998
	    switchport access vlan 998
	    shutdown
	End

    No vlan $($vlan.Substring(1, $vlan.Length - 1))
"@
    }
    $props = @{
        Switch = $hostname
        Interface = $interface
        CurrentConfig = $currentConfig
        Commands = $dTemp
    }
    $return = new-object psobject -Property $props

    return $return
}



function get-EM7Device{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([String])]
    Param
    (
        # Switch Type: Accepts multiple values, limited to: Access, Core, DIST, SD, SDDist
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("Access","Core","DIST","SD","SDDist")]
        [Alias("Type")] 
        $SwitchType,

        # Market Parameter: 
        [Parameter(ParameterSetName='Parameter Set 1')]
        [AllowNull()]
        [AllowEmptyCollection()]
        [AllowEmptyString()]
        [ValidateScript({$true})]
        [Alias("mkt","mktcode")] 
        $Market,

        # EM7 Device ID
        [Parameter(ParameterSetName='EM7 Device ID')]
        [ValidatePattern("[\d\d\d]*")]
        [ValidateLength(3,6)]
        $EM7_ID,
        [Parameter(Mandatory=$false, Position=4)]
        $PortType
    )

    Begin
    {
        #EM7 URI
        $UriPre = 'https://overlook.peak10.com/api/device/'
        $UriPost = '/interface?limit=1000&extended_fetch=1'
        #Credentials
        if(!$creds){$global:creds = Get-Credential -UserName $env:USERNAME -Message 'Enter Network Credentials'}
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($creds.password)
        $pass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        $ErrorActionPreference = "SilentlyContinue" 
        $counter = 0
        $DIST =  ($mkt | where {$_.market -in $market}).dist
        $SDDist = ($mkt | where {$_.market -in $market}).sddist
        $slist_filter02 = $switches | where {$_.name -notlike "*DIST-02"}
        $slist_Market = $slist_filter02 | Where {$_.market -in $market -or $_.name -in $SDDist -or $_.name -in $DIST }
        $slist_type = $slist_Market | where {$_.type -in $switchtype} | Sort-Object type, name
        
        $slist = $slist_type | select Market, Type, Name, EM7_ID
        $totalcount = $slist.count
        $counter = 0
    }
    Process
    {
        write-debug $slist.name

        $EM7Result = foreach($sw in $slist){
            write-debug $sw.name 
            write-debug $counter
            $em7_id = $sw.em7_id
            $uri = "$($UriPre)$($em7_id)$($UriPost)"
            Write-Progress -Activity 'Getting Interfaces' -CurrentOperation $sw.name -PercentComplete (($counter / $totalcount) * 100) -Status "#$($counter) Total: $($totalcount)"
            $Results = Invoke-RestMethod -Method Get -Uri $Uri -Credential $creds
            $ints = $results.result_set |  GM -MemberType NoteProperty | select name
            foreach($int in $ints) {
               if(!$SearchValue)
               {
                    $results.result_set.$($int.name) | select @{N="Market";E={@($sw.Market)}}, @{N="Hostname";E={@($sw.Name)}},@{N="Type";E={@($sw.Type)}}, name, alias, @{N="Sts";E={@(Switch($_.ifoperstatus){1{"Up"}2{ "Down" }})}},  @{N="AdminStatus";E={@(Switch($_.ifAdminStatus){1{"Up"}2{ "Down" }})}}
               }else
               {
                    $results.result_set.$($int.name) | select @{N="Market";E={@($sw.Market)}}, @{N="Hostname";E={@($sw.Name)}},@{N="Type";E={@($sw.Type)}}, name, alias, @{N="Sts";E={@(Switch($_.ifoperstatus){1{"Up"}2{ "Down" }})}},  @{N="AdminStatus";E={@(Switch($_.ifAdminStatus){1{"Up"}2{ "Down" }})}}| where-object {$_.alias -like "*$($searchvalue)*"} 
               }
            }$counter ++
        }
        $EM7Result |ForEach-Object{ $_ | Add-Member -MemberType ScriptProperty -name Status -value {[string]"$($this.sts)/$($this.adminstatus)"}}

    #    $return = New-Object psobject -Property $props
        #New-Object psobject -Property $AliasProperties
        $return = $EM7Result| ForEach-Object {
            New-Object psobject -Property $AliasProperties
            $alias = $_.alias.split(":")
            $hname = $_.hostname.split("-")
            $lookup = lookup-custVLAN -custid $alias[1]
            $ipadd = $Lookup | where {$_.vlan -eq "V$($alias[2])"} | select IPaddress
            $AliasProperties = [ordered]@{
            
                hostname = $_.hostname
                name = $_.name
                Prefix = $alias[0]
                CustID= $alias[1]
                VLAN= $alias[2]
                Suffix= $alias[3]
                Status = $_.status 
                market = $hname[0]
                IPaddress = $ipadd
            }
        }
    }
    End
    {
        foreach($r in $return){
            if($r.name -like "*/*"){
                $split =  ($r.NAME).SPLIT("/")
                [int]$intsplit = $split[($Split.COUNT - 1)]
                $r | Add-Member -MemberType NoteProperty -Name Int -Value $intsplit
                $r | add-member -MemberType NoteProperty -Name PortPrefix -value $split[0]
            }
        }
        if($portType -eq 'Interface'){
            $return | where {$_.name -like "*/*"} | sort hostname, portprefix, int             
            }
        if($portType -eq 'SVI'){
            $return | Where {$_.name -like "*vl*"} | sort hostname, portprefix, int
        }
        if(!$portType){
            return $return | sort hostname, portprefix, int | where {$_.custid -ne $null}
        }
    }

}


function get-Interface_Utilization{
    param(
         [Parameter(Mandatory=$True,Position=1,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
         [string]$hostname,
         [Parameter(Mandatory=$True,Position=2,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
         [Alias("Interface")]
         [string]$name,
         [Parameter(Mandatory=$True,Position=3)]
         [int]$seconds
    
    )
    $start = get-interface -hostname $hostname -name $name

    Start-Sleep -Seconds $seconds
    $finish = get-interface -hostname $hostname -name $name

    $packets_in = [decimal]$finish.packets_in - [decimal]$start.packets_in
    $packets_out = [decimal]$finish.packets_out - [decimal]$start.packets_out 
    $time = ($finish.timestamp - $start.timestamp).Seconds
    $in = ([long]$finish.kb_in - [long]$start.kb_In)/$time
    [long]$out = ([long]$finish.kb_out - [long]$start.KB_out)/$time
    [long]$Total_in = ([long]$finish.KB_in - [long]$start.KB_In)
    [long]$Total_Out = ([long]$finish.kb_out - [long]$start.kb_out)
    $a = [long]$Total_in/1mb
    $b = [long]$Total_Out/1mb
    $props = [ordered]@{
        'Hostname' = $hostname;
        'name' = $name;
        'desc' = $start.desc;
        'IP' = $start.ip;
        'PacketsIn' = $packets_in;
        'PacketsOut' = $packets_out;
        'TotalMBDown' = "{0:N2}" -f $a
        'TotalMBUp' = "{0:N2}" -f $b
        'KBpsDown' = "{0:N2}" -f [long]($In);
        'KBpsUp' = "{0:N2}" -f [long]($out);
        'Seconds Observed' = $time;
        'TimeStamp' = get-date
    }
    $result = New-Object psobject -Property $props    
    return $result
}

Function get-interface{
[CmdletBinding()]
param(
     [Parameter(Mandatory=$True,Position=1,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
     [string]$hostname,
     [Parameter(Mandatory=$True,Position=2,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
     [Alias("Interface")]
     [string]$name
)
    $command = "show interface $($name)"
    $result = get-scommand -Hostname $hostname -command $command
    $result_line = $result -split '[\n]' 
    $5mIn_String = '5 minute input rate*'
    $5mOut_String = '5 minute output rate *'
    $desc_String = 'Description: *'
    $ip_String = 'Internet address is *'
    $packs_in = 'packets input'
    $packs_out = 'packets output'
    $bytes_in = 'input, '
    $bytes_out = 'output, '
    $errors = 'input errors, '

    $5m_Input = foreach($line in $result_line){if($line -match $5mIn_String){$line.trimstart()}}
    $5m_Output = foreach($line in $result_line){if($line -match $5mOut_String){$line.trimstart()}}
    $desc = foreach($line in $result_line){if($line -match $desc_String){$line.trimstart()}}
    $ip = foreach($line in $result_line){if($line -match $ip_String){$line.trimstart()}}
    $packets_in = foreach($line in $result_line){if($line -match $packs_in){$line.trimstart()}}
    $packets_out = foreach($line in $result_line){if($line -match $packs_out){$line.trimstart()}}
    $bytes_in = foreach($line in $result_line){if($line -match $bytes_in){$line.trimstart()}}
    $bytes_out = foreach($line in $result_line){if($line -match $bytes_out){$line.trimstart()}}
    $err = foreach($line in $result_line){if($line -match $errors){$line.trimstart()}}
    $props = [ordered]@{
        'Hostname' = $hostname
        'Interface' = $name
        'desc'  =  $desc.split(" ")[1]
        '5m_Input_KBps'  =  "{0:N0}" -f [long]($5m_Input.split(" ")[4])
        '5m_Output_KBps'  =  "{0:N0}" -f [long]($5m_Output.split(" ")[4])
        'packets_in'  =  "{0:N0}" -f [long]($packets_in.split(" ")[0])
        'packets_out'  =  "{0:N0}" -f [long]($packets_out.split(" ")[0])
        'KB_in'  =  "{0:N2}" -f [long]($packets_in.split(" ")[3]/1KB)
        'KB_out'  =  "{0:N2}" -f [long]($packets_out.split(" ")[3]/1KB)
        'Errors' = $err
    }
    
    $r1 = New-Object psobject -Property $props
    if(!$ip){$r1 | Add-Member -MemberType NoteProperty -Name IP -Value $ip}
    return $r1
}

function Add-InterfaceIndex{
    param($inputobject)
    $index = 0
    $inputobject | Foreach-Object {[PSCustomObject] @{ Index = $index; Hostname = $_.hostname; Name = $_.name}; $index++}
}

function get-VLAN{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        [Parameter(Mandatory=$True,ValueFromPipelineByPropertyName=$true,Position=1)]
        $Hostname,
        [Parameter(Mandatory=$false,Position=2)]
        [string]$Range

    )

    Begin
    {
        #$host.privatedata.ProgressForegroundColor = "yellow"
        #$host.privatedata.ProgressBackgroundColor = "black";
        $counter = 0
        $ScriptStart1 = (Get-Date)
        $totalcount = $hostname.count
        if(!$Range)
        {
            $VLANtypes = import-csv "S:\Provisioning\Central Provisioning\Automation\Inventory\VLAN_Ranges.csv"
            $VLAN = $vlantypes | Out-GridView -PassThru
            $range = $VLAN.range 
            $command = "show vlan id $($range) | include :"
            Write-Debug $command
            if([regex]::Matches($vlan.range,'\b\d{3,4}[-]\d{3,4}').success -eq $true)
            {
                #write-host ($($VLAN.range.split("-")[0])..$($VLAN.range.split("-")[1]))
                $rng = ($($VLAN.range.split("-")[0])..$($VLAN.range.split("-")[1]))
                $RangeDisplay = $VLAN.range
            }
        }else{
            $command = "show vlan id $($Range) | include :"
            Write-Debug $command
            if([regex]::Matches($range, '\b\d{3,4}[-]\d{3,4}').success -eq $true){
                $rng = ($($range.split("-")[0])..$($range.split("-")[1]))
                $RangeDisplay = $range
            }
         }
    }
    Process
    {
        foreach($h in $hostname){
            Write-Verbose $h
            Write-Progress -Activity 'Collecting VLANs' -CurrentOperation $h -PercentComplete (($counter / $totalcount) * 100) -Status "#$($counter) Total: $($totalcount)"
            $return = Get-sCommand -Hostname $h -command $command    
            $return1 = ($return -split '[\r\n]') |? {$_}
            $details = foreach($i in $return1){
                $split0 = $i.split(" ",[System.StringSplitOptions]::RemoveEmptyEntries)
                $vlan = $split0[0]
                $desc = $split0[1]
                $status = $split0[2]
                $custid = $desc.split(":")[1]
                $market= $h.split("-")[0]
                $networkType = $desc.split(":")[0]
                $suffix = $desc.split(":")[3]
                $detail_props = [ordered]@{
                    Market = $market
                    VLAN = $VLAN
                    CustID = $custid
                    Desc = $desc
                    Status = $Status
                    NetType = $networkType
                    Suffix = $suffix
                }
                $obj = new-object psobject -Property $detail_props
                $obj
            }
            if($h -like "*-DIST-*"){
                $rset = Get-sCommand -Hostname $h -command "sh ip interface brief"
                $rset1 = ($rset -split '[\r\n]') |? {$_}
                $rset2 = $rset1 | where {$_ -like "V*"}
                $Int_Rset = foreach($r in $rset2){
                    $split = $r.split(" ",[System.StringSplitOptions]::RemoveEmptyEntries)
                    $Line_Int = $split[0]
                    $line_IP = $split[1]
                    $Int_Chars = $line_int.Length - 4
                    $Line_VLANID = $Line_int.Substring(4,$Int_Chars)
                    $line_props = [ordered]@{
                        VLAN = $Line_VLANID   
                        Interface = $line_int
                        IPAddress = $line_IP
                        Hostname = $hostname
                    }
                $Line_Data = New-Object psobject -Property $line_props
                $line_data  
                }
            }ELSE{
                write-verbose 'no query against sddist Switch Virtual Interfaces'
            }
              
            $Combined_Rset = foreach($d in $details){
                if(!$Int_Rset){
                    $SVI = ''
                    $IPAddress = ''
                }ELSE{
                    $Int_Line = $Int_Rset | where {$_.vlan -eq $d.vlan} 
                    $SVI = $Int_line.Interface
                    $IPaddress = $Int_line.IPAddress
                    $props = [ordered]@{
                            Market = $d.market
                            VLAN = $d.VLAN
                            CustID = $d.custid
                            Desc = $d.desc
                            Status = $d.Status
                            NetType = $d.netType
                            Suffix = $d.suffix
                            IPAddress = $IPaddress
                            SVI = $SVI 
                            Hostname = $h
                            }
                $cRset = New-Object psobject -Property $props
                $cRset
            } 
            }
        $counter ++
        $Combined_Rset
        }
    }
    End
    {
        $ScriptEnd1 = (Get-Date)
        $RunTime1 = New-Timespan -Start $ScriptStart1 -End $ScriptEnd1
        write-Verbose (“Execution Time for Dist: {0}m:{1}s” -f [math]::abs($Runtime1.Minutes),[math]::abs($RunTime1.Seconds)) 
    }
}
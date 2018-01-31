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
        $switches | where {$_.parentmarket -eq $m1} | select -unique name, em7_id, type
    } 
    if(!$SwitchType){$slist = $list}else{
        $slist = $list | where {$_.type -in $SwitchType}
    }
    $slist = foreach($m in $market){
        $m1 = $m.Substring(0,3)
        $switches | where {$_.parentmarket -eq $m1} | select -unique name, em7_id, type
    } 
    #$slist |out-file C:\users\frank.scherer\Documents\switchout.csv

    #EM7 URI
    $UriPre = 'https://overlook.peak10.com/api/device/'
    $UriPost = '/interface?limit=1000&extended_fetch=1'
    #Credentials
    if(!$creds){$global:creds = Get-Credential -UserName $env:USERNAME -Message 'Enter EM7 Password'}
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($creds.password)
    $pass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    $ErrorActionPreference = "SilentlyContinue"


    $counter = 0    

    #$slist = (get-marketDevices -market $market) | where {$_.Type -ne 'Core' -and $_.name -notlike '*dist-02'}
    #write-debug $slist.name
    $totalcount = $slist.count

    $EM7Result = foreach($sw in $slist){
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
        $AliasProperties = [ordered]@{
            
            hostname = $_.hostname
            name = $_.name
            Prefix = $alias[0]
            CustID= $alias[1]
            VLAN= $alias[2]
            Suffix= $alias[3]
            Status = $_.status 
            market = $hname[0]
            IPaddress = ''
        }
    }
    
    $return | Add-Member -MemberType ScriptMethod -Name ShowRun -Value {
            $result = Get-sCommand -Hostname $this.hostname  -command "show run interface $($this.name)"
            $this | add-member -membertype NoteProperty -name ShowRun -value $result
            return $result
    }

    $return | Add-member -MemberType ScriptMethod -Name GetIntIP -value force {
        $dists = $this | select -unique hostname |where {$_.hostname -like "*-DIST-01"}
        $global:intIP = foreach($d in $dists){
            $rset = Get-sCommand -Hostname $d.hostname -command "sh ip interface brief"
            $r1 = ($rset -split '[\r\n]') |? {$_}
            $distVLANs = $this | select hostname, vlan, name | where {$_.name -like "vl*" -and $_.hostname -like "*-DIST-01"}
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
    return $return
}
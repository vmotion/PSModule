function get-EM7Device
{
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
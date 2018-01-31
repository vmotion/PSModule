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
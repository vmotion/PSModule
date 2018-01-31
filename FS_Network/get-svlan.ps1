function get-sVLAN
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        [Parameter(Mandatory=$True,ValueFromPipelineByPropertyName=$true,Position=1)]
        $Hostname,
        [Parameter(Mandatory=$false,Position=2)]
        $Range

    )

    Begin
    {
        $counter = 0
        $ScriptStart1 = (Get-Date)
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
            Write-Progress -Activity 'Connecting to $h' -CurrentOperation $h -PercentComplete (($counter / $hostname.count) * 100)
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
                $obj | Add-member -MemberType NoteProperty -Name Market -value $market 
                $obj | Add-member -MemberType NoteProperty -Name VLAN -value $VLAN
                $obj | Add-member -MemberType NoteProperty -Name CustID -Value $custid
                $obj | Add-member -MemberType NoteProperty -Name Desc -value $desc
                $obj | Add-member -MemberType NoteProperty -Name Status -value $Status
                $obj | add-member -MemberType NoteProperty -Name NetType -value $networkType 
                $obj | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                $defaultProperties = @('Market','VLAN','CustID','NetType')
                $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultProperties)
                $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
                $counter ++
                $obj
            }
            $details
        }
    }
    End
    {
        $ScriptEnd1 = (Get-Date)
        $RunTime1 = New-Timespan -Start $ScriptStart1 -End $ScriptEnd1
        write-Verbose (“Execution Time for Dist: {0}m:{1}s” -f [math]::abs($Runtime1.Minutes),[math]::abs($RunTime1.Seconds)) -ForegroundColor Cyan
    }
}
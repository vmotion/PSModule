function update-vlandata {
    $mkt = $switches | Select-Object -Unique market
    $slist = $switches | Where-Object {$_.name -like "*DIST-*"} | Select-Object -Unique name

    $date = get-date -Format {MM/dd/yyyy}
    $time = get-date -format {HH:mm}
    $data = get-VLAN -Hostname $slist.name -Range '200-1799' -Verbose


    $data | export-csv 'S:\Provisioning\Central Provisioning\Automation\Inventory\VLANs.CSV'
    $new = $data | Group-Object Market | Select-Object name, @{N="Count_Current";E={@($_.count)}}| Sort-Object -Descending count 
    $sum = import-csv 'S:\Provisioning\Central Provisioning\Automation\Inventory\VLAN_Summary.csv'
    $summary = $sum | Select-Object -Property * -ExcludeProperty difference
    $summary | Add-Member -MemberType NoteProperty -Name $($date) -Value '' 
    foreach($i in $summary){
       $value = ($new | where {$_.name -eq $i.Market}).count_current
       $i.$date = $value
    }

    $lastUpdate = ($summary | gm -type NoteProperty).name |where {$_ -ne 'Market'} |Sort-Object | select -First 1
    $summary | ForEach-Object{$_ | Add-Member -force -MemberType ScriptProperty -Name Difference -Value {($this.$date) - ($this.$lastupdate)}}
    $result = $summary | Select-Object Market, $lastUpdate, $date, difference | Sort-Object difference, market
    $result | export-csv 'S:\Provisioning\Central Provisioning\Automation\Inventory\VLAN_Summary.csv'
    return $result
    write-host "File Updated: S:\Provisioning\Central Provisioning\Automation\Inventory\VLANs.CSV"

}


function get-scriptroot{
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


return $config
}
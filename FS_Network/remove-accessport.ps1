function Remove-AccessPort{
        [CmdletBinding()]
    param ( 
        [Parameter(Mandatory=$True,Position=1)] 
        [string]$hostname, 
        [Parameter(Mandatory=$True,Position=2)]
        [string]$interface,
        [Parameter(Mandatory=$True,Position=3)]
        [string]$vlan
    )

$currentConfig = Get-sIntConfig -hostname $hostname -interface $interface 
if($interface -like "VL*"){
    $dtemp = @"
        No interface $($vlan.Substring(1, $vlan.Length - 1))
"@
}else{
$dTemp = @"
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


$return = New-Object psobject -Property $props

}
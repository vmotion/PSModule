function get-VM_Networks {
Param
    (

        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $custid
                   
    )
write-host $custid
$folderName = $custid
$tgtFolder = Get-Folder -Name $folderName


$reportedvms=New-Object System.Collections.ArrayList
$vms=Get-View -ViewType VirtualMachine -Filter @{"Parent"=$tgtFolder.ExtensionData.MoRef.Value} |Sort-Object -Property {  $_.Config.Hardware.Device |  where {$_ -is [VMware.Vim.VirtualEthernetCard]} |  Measure-Object | select -ExpandProperty Count} -Descending
 
    foreach($vm in $vms){
        $reportedvm = New-Object PSObject
        Add-Member -Inputobject $reportedvm -MemberType noteProperty -name Guest -value $vm.Name
        Add-Member -InputObject $reportedvm -MemberType noteProperty -name UUID -value $($vm.Config.Uuid)
        Add-Member -InputObject $reportedvm -MemberType NoteProperty -name Hostname -value $vm.Guest.HostName
        $networkcards=$vm.guest.net | ?{$_.DeviceConfigId -ne -1}
        $i=0
        foreach($ntwkcard in $networkcards){
            Add-Member -InputObject $reportedvm -MemberType NoteProperty -Name "networkcard${i}.Network" -Value $ntwkcard.Network
            Add-Member -InputObject $reportedvm -MemberType NoteProperty -Name "networkcard${i}.MacAddress" -Value $ntwkcard.Macaddress  
            Add-Member -InputObject $reportedvm -MemberType NoteProperty -Name "networkcard${i}.IpAddress" -Value $($ntwkcard.IpAddress|?{$_ -like "*.*"})
            Add-Member -InputObject $reportedvm -MemberType NoteProperty -Name "networkcard${i}.Device" -Value $(($vm.config.hardware.device|?{$_.key -eq $($ntwkcard.DeviceConfigId)}).gettype().name)
            add-member -InputObject $reportedvm -MemberType NoteProperty -name 'Network' -value 
            $i++
        }
        $reportedvms.add($reportedvm)|Out-Null
    }
} 
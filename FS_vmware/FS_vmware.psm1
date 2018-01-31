<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   General notes
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>
function connect-vcenter
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([String])]
    Param
    (
        # function to connect to a peak 10 vCenter
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [ValidateCount(0,5)]
        [ValidateSet("ATL","ATL2","CLT","LOU","NAS","Peak1005")]
        [Alias("mkt")] 
        $Market
    )

    Begin
    {
        get-module *vmware* -ListAvailable | Import-Module
        if((test-path 'S:\Provisioning\CIE_Notebook\InvLists') -eq $true){
                $vcs = Import-Csv 'S:\Provisioning\CIE_Notebook\InvLists\vCenter_Servers.csv'
            }Else{
                $vcs = Import-Csv 'C:\scripts\Inventory\vCenter_Servers.csv'
            }
    
        $vCenterServer = ($vcs | Where-Object {$_.market -eq $market}).name
        $currentuser = $env:USERNAME
        $SSODomain = 'peak10ms.com'
        $vcuser = $SSODomain + '\' + $currentuser
        $creds = Get-Credential -message "Enter password for $vcuser on $vCenterServer" -user $vcuser 
    }
    Process{
        if ($pscmdlet.ShouldProcess("Target", "Operation")){
                Connect-VIServer -Server $vcenterServer -Credential $creds -WarningAction SilentlyContinue -ErrorAction Stop | Out-Null
        }
    }
    End
    {
    }
}

function Find-VPG{
    param(
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $VPGName        
    )
    $selected_hosts = get-vmhost | Sort-Object Name | Out-GridView -PassThru
    write-host "You selected $($selected_hosts.count) Hosts" -ForegroundColor Yellow  
    write-host 'Getting Virtual PortGroups for all selected hosts. This may take a moment..' -ForegroundColor Yellow
    $VPGs = foreach($shost in $selected_hosts){
        Get-VirtualPortGroup -VMHost $shost | select @{Expression={$shost};Label="Host"}, Name, VLANid, VirtualSwitchName
    }
    #$VPGs | Export-Csv 'S:\Provisioning\Central Provisioning\Automation\Inventory\VPG_Export.csv'
    if(!$VPGName) {$VPGName= ($VPGs |select -unique Name, VLANid, VirtualSwitchName | Out-GridView -Title 'Select a VPG.' -PassThru).name}
    $VPGResult = @() 
        foreach ($shost in $selected_hosts) {
        $hostVPGs = get-virtualportgroup -vmhost $shost -Name $vpgname -ErrorAction SilentlyContinue| select @{Expression={$shost};Label="Host"}, Name, VLANid, VirtualSwitchName 
        $VPGresult += $hostVPGs
     }

write-host "$($vpgname) was found on $($VPGResult.count) of $($selected_hosts.count) Hosts" -ForegroundColor Yellow 
return $VPGResult
}

function Find-CustomerVPG{
    param(
        [Parameter(Mandatory=$True,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Customer        
    )
    $selected_hosts = get-vmhost | Sort-Object Name | Out-GridView -PassThru
    write-host "You selected $($selected_hosts.count) Hosts" -ForegroundColor Yellow  
    write-host 'Getting Virtual PortGroups for all selected hosts. This may take a moment..' -ForegroundColor Yellow
    $customersearch = "*$($customer)*"
    $VPGResult = @() 
        foreach ($shost in $selected_hosts) {
        $hostVPGs = get-virtualportgroup -vmhost $shost -ErrorAction SilentlyContinue| select @{Expression={$shost};Label="Host"}, Name, VLANid, VirtualSwitchName| where-object {$_.name -like $customersearch}  
        $VPGresult += $hostVPGs
     }
return $vpgresult
}

function Get-CustomerVM {
$location = Read-Host -Prompt 'Input the VM Folder' 
$docs = "$env:USERPROFILE\documents"
$csvout = "$location - VMs.csv"
$vm = Get-VM -location $location | select Name, @{N="Hostname";E={@($_.guest.hostname)}}, PowerState, Notes, Guest, NumCpu, CoresPerSocket, MemoryMB, MemoryGB, UsedSpaceGB, ProvisionedSpaceGB, @{N="Internal IP";E={@($_.guest.IPAddress[0])}}, @{N="SDNet IP";E={@($_.guest.IPAddress[2])}}
write-host "Building Variable List..."
$vm | Export-csv $csvout
Invoke-Item $csvout
}

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
            $i++
        }
        $reportedvms.add($reportedvm)|Out-Null
    }
return $reportedvms
} 


function new-sVirtualPortGroup{
    [CmdletBinding()]
    param ( 
        [Parameter(Mandatory=$True,Position=1)] 
        $VPGName, 
        [Parameter(Mandatory=$true,Position=2)]
        $VLAN
        )
    Get-Module -ListAvailable '*vmware*' | Import-Module
    $currentuser = $env:USERNAME
    $SSODomain = 'peak10ms.com'
    $vcuser = $SSODomain + '\' + $currentuser
    if(!$creds) {$creds = Get-Credential -UserName $vcuser -Message 'Enter Peak10MS.com Password'}
    Connect-VIServer -Server 'PEAK1005-VCS98.peak10ms.com' -Credential $creds
    $cluster = get-cluster | Out-GridView -PassThru
    $hosts = foreach($c in $cluster) { $c | Get-VMHost }
    $switches = $hosts | Get-VirtualSwitch |Out-GridView -PassThru
    $tgtFolder = get-folder -type Network | Out-GridView -PassThru -Title 'Select the Folder'
    foreach($i in $switches){
        New-VirtualPortGroup -Name $vpgname -VirtualSwitch $i -VLanId $VLAN
    }

    # Get network folder
    $esx = $hosts | Get-View
    $dc = Get-Datacenter -VMHost $hosts | Get-View
    $netFolder = Get-View $dc.NetworkFolder
    $custid =  "*$($vpgname.split("-")[0])*"
    $value = "$($tgtFolder.Id.split("-")[1])-$($tgtFolder.Id.split("-")[2])"

    $networks = $netFolder.ChildEntity | where {$_.Type -eq 'Network'} | %{ get-view $_ } | where {$_.name -like "*$($custid)*"} 
    $folder = $tgtFolder | %{ get-view $_ }

    foreach($net in $networks){
        $pgMoRef = $net.moref 
 
        $folder | %{
            $child = $_
            if($net.Name.split("-")[0] -eq $net.name.split("-")[0]){
                $child.MoveIntoFolder($pgMoRef)
            }
        }
    }

    $analysis = $hosts | select name, id
    $switchlookup = Get-VirtualPortGroup -name $VPGNAME | select name, VMHostId 
    $analysis | ForEach-Object {$_| Add-Member -MemberType ScriptProperty -Name 'VPG' -Value {($switchlookup | where {$_.vmhostid -eq $this.id}).name}}
    $return = $analysis | select @{N="HostName";E={@($_.name.split(".")[0])}},  vpg
    return $return | sort hostname

}


function Move-sVirtualPortGroup{
    [CmdletBinding()]
    param ( 
        [Parameter(Mandatory=$True,Position=1)] 
        $VPGName, 
        [Parameter(Mandatory=$true,Position=2)]
        $VLAN
        )
    Get-Module -ListAvailable '*vmware*' | Import-Module
    $currentuser = $env:USERNAME
    $SSODomain = 'peak10ms.com'
    $vcuser = $SSODomain + '\' + $currentuser
    if(!$creds) {$creds = Get-Credential -UserName $vcuser -Message 'Enter Peak10MS.com Password'}
    Connect-VIServer -Server 'PEAK1005-VCS98.peak10ms.com' -Credential $creds
    $cluster = get-cluster | Out-GridView -PassThru
    $hosts = foreach($c in $cluster) { $c | Get-VMHost }
    $switches = $hosts | Get-VirtualSwitch |Out-GridView -PassThru
    $tgtFolder = get-folder -type Network | Out-GridView -PassThru -Title 'Select the Folder'
    # Get network folder
    $esx = $hosts | Get-View
    $dc = Get-Datacenter -VMHost $hosts | Get-View
    $netFolder = Get-View $dc.NetworkFolder
    $custid =  "*$($vpgname.split("-")[0])*"
    $value = "$($tgtFolder.Id.split("-")[1])-$($tgtFolder.Id.split("-")[2])"

    $networks = $netFolder.ChildEntity | where {$_.Type -eq 'Network'} | %{ get-view $_ } | where {$_.name -like "*$($custid)*"} 
    $folder = $tgtFolder | %{ get-view $_ }

    foreach($net in $networks){
        $pgMoRef = $net.moref 
 
        $folder | %{
            $child = $_
            if($net.Name.split("-")[0] -eq $net.name.split("-")[0]){
                $child.MoveIntoFolder($pgMoRef)
            }
        }
    }

    $analysis = $hosts | select name, id
    $switchlookup = Get-VirtualPortGroup -name $VPGNAME | select name, VMHostId 
    $analysis | ForEach-Object {$_| Add-Member -MemberType ScriptProperty -Name 'VPG' -Value {($switchlookup | where {$_.vmhostid -eq $this.id}).name}}
    $return = $analysis | select @{N="HostName";E={@($_.name.split(".")[0])}},  vpg
    return $return | sort hostname

}
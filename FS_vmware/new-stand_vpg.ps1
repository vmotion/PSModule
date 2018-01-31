function new-sVirtualPortGroup{
    [CmdletBinding()]
    param ( 
        [Parameter(Mandatory=$True,Position=1,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)] 
        $VPGName, 
        [Parameter(Mandatory=$false,Position=2)]
        [string]$VLAN,
        [Parameter(mandatory=$false,Position=3,ValueFromPipelineByPropertyName=$true)]
        $SwitchType
        )
    Get-Module -ListAvailable '*vmware*' | Import-Module
    $currentuser = $env:USERNAME
    $SSODomain = 'peak10ms.com'
    $vcuser = $SSODomain + '\' + $currentuser
    if(!$creds) {$creds = Get-Credential -UserName $vcuser -Message 'Enter Peak10MS.com Password'}
    Connect-VIServer -Server 'PEAK1005-VCS98.peak10ms.com' -Credential $creds
    $cluster = get-cluster | Out-GridView -PassThru
    $switches = $hosts | Get-VirtualSwitch |Out-GridView -PassThru
    $tgtFolder = get-folder -type Network | Out-GridView -PassThru -Title 'Select the Folder'
    foreach($i in $switches){
        New-VirtualPortGroup -Name $vpgname -VirtualSwitch $i -VLanId $vlan
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
function connect-vcenter{
<# 
  .SYNOPSIS  
    Connect Posh Session to vCenter Server 
  .EXAMPLE 
    Connect-vCenter -Market CLT

#>

    $vcs = Import-Csv .\vCenter_Servers.csv
    $vCenterServer = ($vcs | Where-Object {$_.market -eq $market}).name
    $currentuser = $env:USERNAME
    $SSODomain = 'peak10ms.com'
    $vcuser = $SSODomain + '\' + $currentuser
    $creds = Get-Credential -message "Enter password for $vcuser on $vCenterServer" -user $vcuser

    Connect-VIServer -Server $vcenterServer -Credential $creds -WarningAction SilentlyContinue
    $invpath = 'C:\Scripts\Inventory'
}





function Get-VLAN_PO{
    [CmdletBinding()]
    param ( 
        [Parameter(Mandatory=$false,Position=1)]
        [string]$VLANID,
        [parameter(Mandatory=$false,position=2)]
        [System.Object]$hostname
        )

        $marketsplit = $hostname.split("-")[0]
        $aswitches = $switches | where {$_.type -eq 'access' -and $_.name -like "*$($marketsplit)*"}
#Port-Channel
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($creds.password)
            $pass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
            $result = Get-sCommand -Hostname $hostname -command "show vlan id $($vlanid)"
            $POValues = Select-String '(Po\d{1,4})' -input $result -AllMatches | Foreach {$_.matches} | select @{N="PortChannel";E={@($_.value)}}, @{N="S1";E={@($_.value.length-2)}}
            $POValues | ForEach-Object {$_ | add-member -MemberType NoteProperty -name 'SwitchNum' -value $_.portchannel.substring($_.S1,2) }
            $POValues | add-member -MemberType ScriptProperty -name 'Switch' -value (($switches |Where-Object {$_.po -eq $this.PortChannel -and $_.market -like "$($marketsplit)*"}).name)
            #write-host $result
            $VLANPO = new-object PSobject
            $vlanpo | add-member -MemberType NoteProperty -Name 'VLAN' -Value $VLANID
            $vlanpo | add-member -MemberType NoteProperty -Name 'PortChannels' -value $POValues.portchannel

            return $PoValues
            


}
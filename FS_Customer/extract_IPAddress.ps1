        $intIP = foreach($d in $dists){
            $rset = Get-sCommand -Hostname $d.hostname -command "sh ip interface brief"
            $r1 = ($rset -split '[\r\n]') |? {$_}
            $distVLANs = $INTEM001_Ints | select hostname, vlan, name | where {$_.name -like "vl*" -and $_.hostname -like "*-DIST-0*"}
            $vl = $distvlans | select name,vlan, @{N="VLANID";E={@(($_.vlan).substring(1,($_.vlan).Length - 1) )}}
            $regex =  '\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b'
            $IP_Lookup = foreach($v in $vl){
                $z = $r1 | where {$_ -like "Vlan$($v.vlanid)*"}
                $split = $z.split(" ",[System.StringSplitOptions]::RemoveEmptyEntries)
                $vname = $split[0]
                $ip = $split[1]
                $props = @{
                    vName = $vname
                    IPAddress = $ip
                }
                $IP_Lookupet = New-Object psobject -Property $props
                $IP_Lookupet  
            }
            $IP_Lookup | select @{N="VLAN";E={@("V"+$_.vName.substring(4,(($_.vname).length - 4)))}}, IPaddress
        }
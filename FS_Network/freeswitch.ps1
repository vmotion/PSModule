$feh | add-member -MemberType ScriptMethod -Name FreeSwitch -Value {         
    param([string]$parameter01 = $(throw "Must supply a switch number."))
    $p = "*$($parameter01)*"         
    $a = $clt.freeports | where {$_.hostname -like $p} | Sort-Object|  %{[int]$_.name.split("/")[1]}         
    $b = $clt.freeports | select -Unique hostname | where {$_.hostname -like $p} 
    write-host $b.hostname -ForegroundColor Yellow 
    return $a
} -force

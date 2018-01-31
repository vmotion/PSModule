Function get-interface_stats{
param(
     [Parameter(Mandatory=$True,Position=1,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
     $hostname,
     [Parameter(Mandatory=$True,Position=2,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
     $interface
)
    $command = "show interface $($Interface)"
    $result = get-scommand -Hostname $hostname -command $command
    $result_line = $result -split '[\n]' 

    $packs_in = [regex]::Match($result_line, '\d+ packets input()')
    $packs_out = [regex]::Match($result_line, '\d+ packets output()')
    $bytes_in = [regex]::Match($result_line, 'input, [\d]+() bytes')
    $bytes_out = [regex]::Match($result_line, 'output, [\d]+() bytes')
    $a =  (($bytes_in.value).split(" "))[1]/1MB
    $b = (($bytes_out.value).split(" "))[1]/1MB

    $props = [ordered]@{
        'Hostname' = $hostname;
        'Interface' = $interface;
        'Pkts_In' = "{0:N0}" -f [long](($packs_in.value).split(" "))[0];  
        'Pkts_Out' = "{0:N0}" -f [long](($packs_out.value).split(" "))[0];
        'MB_In' = "{0:N2}" -f $a;
        'MB_out' = "{0:N2}" -f $b
        'TimeStamp' =get-date
    }

    $stats = new-object psobject -Property $props
    return $stats
}
Function get-interface_stats{
param(
     [Parameter(Mandatory=$True,Position=1,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
     $hostname,
     [Parameter(Mandatory=$True,Position=2,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
     [Alias("Interface")]
     $name
)
    $command = "show interface $($name)"
    $result = get-scommand -Hostname $hostname -command $command
    $result_line = $result -split '[\n]' 

    $packs_in = [regex]::Match($result_line, '\d+ packets input()')
    $packs_out = [regex]::Match($result_line, '\d+ packets output()')
    $bytes_in = [regex]::Match($result_line, 'input, [\d]+() bytes')
    $bytes_out = [regex]::Match($result_line, 'output, [\d]+() bytes')
    $a =  [long]((($bytes_in.value).split(" "))[1])/1KB
    $b = [long]((($bytes_out.value).split(" "))[1])/1KB

    $props = [ordered]@{
        'Hostname' = $hostname;
        'Name' = $name;
        'Pkts_In' = "{0:N0}" -f [long](($packs_in.value).split(" "))[0];  
        'Pkts_Out' = "{0:N0}" -f [long](($packs_out.value).split(" "))[0];
        'KB_In' = "{0:N2}" -f $a;
        'KB_out' = "{0:N2}" -f $b
        'TimeStamp' =get-date
    }

    $stats = new-object psobject -Property $props
    return $stats
}
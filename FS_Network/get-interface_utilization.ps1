function get-Interface_Utilization{
    param(
         [Parameter(Mandatory=$True,Position=1,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
         [string]$hostname,
         [Parameter(Mandatory=$True,Position=2,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
         [Alias("Interface")]
         [string]$name,
         [Parameter(Mandatory=$True,Position=3)]
         [int]$seconds
    
    )
    $start = get-interface_stats -hostname $hostname -name $name

    Start-Sleep -Seconds $seconds
    $finish = get-interface_stats -hostname $hostname -name $name

    $packets_in = [decimal]$finish.pkts_in - [decimal]$start.Pkts_In
    $packets_out = [decimal]$finish.pkts_out - [decimal]$start.Pkts_Out 
    $time = ($finish.timestamp - $start.timestamp).Seconds
    $in = ([long]$finish.kb_in - [long]$start.kb_In)/$time
    [long]$out = ([long]$finish.kb_out - [long]$start.KB_out)/$time
    [long]$Total_in = ([long]$finish.KB_in - [long]$start.KB_In)
    [long]$Total_Out = ([long]$finish.kb_out - [long]$start.kb_out)
    $a = [long]$Total_in/1mb
    $b = [long]$Total_Out/1mb
    $props = [ordered]@{
        'Hostname' = $hostname;
        'name' = $name;
        'PacketsIn' = $packets_in;
        'PacketsOut' = $packets_out;
        'TotalMBDown' = "{0:N2}" -f $a
        'TotalMBUp' = "{0:N2}" -f $b
        'KBpsDown' = "{0:N2}" -f [long]($In);
        'KBpsUp' = "{0:N2}" -f [long]($out);
        'Seconds Observed' = $time

    }
    $result = New-Object psobject -Property $props    
    return $result
}
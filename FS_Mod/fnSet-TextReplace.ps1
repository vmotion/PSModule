function Set-textReplace{
    param ( 
        [Parameter(Mandatory=$True,Position=1)] 
        [string]$inpath, 
        [Parameter(Mandatory=$True,Position=2)]
        [string]$outpath,
        [parameter(Mandatory=$true,Position=3)]
        [psobject]$inputObject

        ) 
    
    $result = get-content $inpath | ForEach-Object {
        $line = $_
        $inputObject.getenumerator()|ForEach-Object {
            if($line -match $_.key){
                $line = $line -replace $_.tag, $_.text
            }
        }
    $Line
    }|set-content -path $outpath
}


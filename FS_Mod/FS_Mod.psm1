function Get-Subnet {
    [CmdletBinding()] 
    param ( 
        [Parameter(Mandatory=$True,Position=1)] 
        [string]$IPAddress, 
        [Parameter(Mandatory=$False,Position=2)] 
        [string]$Netmask, 
        [switch]$IncludeTextOutput, 
        [switch]$IncludeBinaryOutput 
        ) 
 
    #region HelperFunctions 
 
    # Function to convert IP address string to binary 
    function toBinary ($dottedDecimal){ 
     $dottedDecimal.split(".") | ForEach-Object {$binary=$binary + $([convert]::toString($_,2).padleft(8,"0"))} 
     return $binary 
    } 
 
    # Function to binary IP address to dotted decimal string 
    function toDottedDecimal ($binary){ 
     do {$dottedDecimal += "." + [string]$([convert]::toInt32($binary.substring($i,8),2)); $i+=8 } while ($i -le 24) 
     return $dottedDecimal.substring(1) 
    } 
 
    # Function to convert CIDR format to binary 
    function CidrToBin ($cidr){ 
        if($cidr -le 32){ 
            [Int[]]$array = (1..32) 
            for($i=0;$i -lt $array.length;$i++){ 
                if($array[$i] -gt $cidr){$array[$i]="0"}else{$array[$i]="1"} 
            } 
            $cidr =$array -join "" 
        } 
        return $cidr 
    } 
 
    # Function to convert network mask to wildcard format 
    function NetMasktoWildcard ($wildcard) { 
        foreach ($bit in [char[]]$wildcard) { 
            if ($bit -eq "1") { 
                $wildcardmask += "0" 
                } 
            elseif ($bit -eq "0") { 
                $wildcardmask += "1" 
                } 
            } 
        return $wildcardmask 
        } 
    #endregion 
 
 
    # Check to see if the IP Address was entered in CIDR format. 
    if ($IPAddress -like "*/*") { 
        $CIDRIPAddress = $IPAddress 
        $IPAddress = $CIDRIPAddress.Split("/")[0] 
        $cidr = [convert]::ToInt32($CIDRIPAddress.Split("/")[1]) 
        if ($cidr -le 32 -and $cidr -ne 0) { 
            $ipBinary = toBinary $IPAddress 
            Write-Verbose $ipBinary 
            $smBinary = CidrToBin($cidr) 
            Write-Verbose $smBinary 
            $Netmask = toDottedDecimal($smBinary) 
            $wildcardbinary = NetMasktoWildcard ($smBinary) 
            } 
        else { 
            Write-Warning "Subnet Mask is invalid!" 
            Exit 
            } 
        } 
 
    # Address was not entered in CIDR format. 
    else { 
        if (!$Netmask) { 
            $Netmask = Read-Host "Netmask" 
            } 
        $ipBinary = toBinary $IPAddress 
        if ($Netmask -eq "0.0.0.0") { 
            Write-Warning "Subnet Mask is invalid!" 
            Exit 
            } 
        else { 
            $smBinary = toBinary $Netmask 
            $wildcardbinary = NetMasktoWildcard ($smBinary) 
            } 
        } 
 
 
    # First determine the location of the first zero in the subnet mask in binary (if any) 
    $netBits=$smBinary.indexOf("0") 
 
    # If there is a 0 found then the subnet mask is less than 32 (CIDR). 
    if ($netBits -ne -1) { 
        $cidr = $netBits 
        #validate the subnet mask 
        if(($smBinary.length -ne 32) -or ($smBinary.substring($netBits).contains("1") -eq $true)) { 
            Write-Warning "Subnet Mask is invalid!" 
            Exit 
            } 
        # Validate the IP address 
        if($ipBinary.length -ne 32) { 
            Write-Warning "IP Address is invalid!" 
            Exit 
            } 
        #identify subnet boundaries 
        $networkID = toDottedDecimal $($ipBinary.substring(0,$netBits).padright(32,"0")) 
        $networkIDbinary = $ipBinary.substring(0,$netBits).padright(32,"0") 
        $firstAddress = toDottedDecimal $($ipBinary.substring(0,$netBits).padright(31,"0") + "1") 
        $firstAddressBinary = $($ipBinary.substring(0,$netBits).padright(31,"0") + "1") 
        $lastAddress = toDottedDecimal $($ipBinary.substring(0,$netBits).padright(31,"1") + "0") 
        $lastAddressBinary = $($ipBinary.substring(0,$netBits).padright(31,"1") + "0") 
        $broadCast = toDottedDecimal $($ipBinary.substring(0,$netBits).padright(32,"1")) 
        $broadCastbinary = $ipBinary.substring(0,$netBits).padright(32,"1") 
        $wildcard = toDottedDecimal ($wildcardbinary) 
        $Hostspernet = ([convert]::ToInt32($broadCastbinary,2) - [convert]::ToInt32($networkIDbinary,2)) - 1 
       } 
 
    # Subnet mask is 32 (CIDR) 
    else { 
     
        # Validate the IP address 
        if($ipBinary.length -ne 32) { 
            Write-Warning "IP Address is invalid!" 
            Exit 
            } 
 
        #identify subnet boundaries 
        $networkID = toDottedDecimal $($ipBinary) 
        $networkIDbinary = $ipBinary 
        $firstAddress = toDottedDecimal $($ipBinary) 
        $firstAddressBinary = $ipBinary 
        $lastAddress = toDottedDecimal $($ipBinary) 
        $lastAddressBinary = $ipBinary 
        $broadCast = toDottedDecimal $($ipBinary) 
        $broadCastbinary = $ipBinary 
        $wildcard = toDottedDecimal ($wildcardbinary) 
        $Hostspernet = 1 
        $cidr = 32 
        } 
 
    #region Output 
 
    # Include a ipcalc.pl style text output (not an object) 
    if ($IncludeTextOutput) { 
        Write-Host "`nAddress:`t`t$IPAddress" 
        Write-Host "Netmask:`t`t$Netmask = $cidr" 
        Write-Host "Wildcard:`t`t$wildcard" 
        Write-Host "=>" 
        Write-Host "Network:`t`t$networkID/$cidr" 
        Write-Host "Broadcast:`t`t$broadCast" 
        Write-Host "HostMin:`t`t$firstAddress" 
        Write-Host "HostMax:`t`t$lastAddress" 
        Write-Host "Hosts/Net:`t`t$Hostspernet`n" 
        } 
 
    # Output custom object with or without binary information. 
    if ($IncludeBinaryOutput) { 
        [PSCustomObject]@{ 
            Address = $IPAddress 
            Netmask = $Netmask 
            Wildcard = $wildcard 
            Network = "$networkID/$cidr" 
            Broadcast = $broadCast 
            HostMin = $firstAddress 
            HostMax = $lastAddress 
            'Hosts/Net' = $Hostspernet 
            AddressBinary = $ipBinary 
            NetmaskBinary = $smBinary 
            WildcardBinary = $wildcardbinary 
            NetworkBinary = $networkIDbinary 
            HostMinBinary = $firstAddressBinary 
            HostMaxBinary = $lastAddressBinary 
            BroadcastBinary = $broadCastbinary 
            } 
        } 
    else { 
        [PSCustomObject]@{ 
            Address = $IPAddress 
            Netmask = $Netmask 
            Wildcard = $wildcard 
            Network = "$networkID/$cidr" 
            Broadcast = $broadCast 
            HostMin = $firstAddress 
            HostMax = $lastAddress 
            'Hosts/Net' = $Hostspernet 
            }     
        }
 } 
#endregion

function Import-XAML{
<#
.Synopsis
   WPF Import tool for Powershell. WPF is a platform for developing user interfaces. This function will import a WPF
   File to a Powershell Object named WFP. 
.DESCRIPTION
   WPF Import Utility
.EXAMPLE
   import-xaml 'C:\MainWindow.WPF' Form

#>

    [CmdletBinding()]
    [Alias()]
    Param
    (
        # WPF File to import
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $FileName
    )

    Begin
    {
    
        Add-Type -AssemblyName presentationframework, presentationcore


    }
    Process
    {

        $wpf = @{ }
        $inputXML =  Get-Content -Path $FileName
        $inputXMLClean = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace 'x:Class=".*?"','' -replace 'd:DesignHeight="\d*?"','' -replace 'd:DesignWidth="\d*?"',''
        [xml]$xaml = $inputXMLClean
        $reader = New-Object System.Xml.XmlNodeReader $xaml
        $tempform = [Windows.Markup.XamlReader]::Load($reader)
        $Nodes = $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")
        $Nodes | ForEach-Object {$wpf.Add($_.Name, $tempform.FindName($_.Name))}
    }
    End
    {
        Return $wpf
    }
}

Function Get-SubNetItems
{
<# 
	.SYNOPSIS 
		Scan subnet machines
		
	.DESCRIPTION 
		Use Get-SubNetItems to receive list of machines in specific IP range.

	.PARAMETER StartScanIP 
		Specify start of IP range.

	.PARAMETER EndScanIP
		Specify end of IP range.

	.PARAMETER Ports
		Specify ports numbers to scan if open or not.
		
	.PARAMETER MaxJobs
		Specify number of threads to scan.
		
	.PARAMETER ShowAll
		Show even adress is inactive.
	
	.PARAMETER ShowInstantly 
		Show active status of scaned IP address instanly. 
	
	.PARAMETER SleepTime  
		Wait time to check if threads are completed.
 
	.PARAMETER TimeOut 
		Time out when script will be break.

	.EXAMPLE 
		PS C:\>$Result = Get-SubNetItems -StartScanIP 10.10.10.1 -EndScanIP 10.10.10.10 -ShowInstantly -ShowAll
		10.10.10.7 is active.
		10.10.10.10 is active.
		10.10.10.9 is active.
		10.10.10.1 is inactive.
		10.10.10.6 is active.
		10.10.10.4 is active.
		10.10.10.3 is inactive.
		10.10.10.2 is active.
		10.10.10.5 is active.
		10.10.10.8 is inactive.

		PS C:\> $Result | Format-Table IP, Active, WMI, WinRM, Host, OS_Name -AutoSize

		IP           Active   WMI WinRM Host              OS_Name
		--           ------   --- ----- ----              -------
		10.10.10.1    False False False
		10.10.10.2     True  True  True pc02.mydomain.com Microsoft Windows Server 2008 R2 Enterprise
		10.10.10.3    False False False
		10.10.10.4     True  True  True pc05.mydomain.com Microsoft Windows Server 2008 R2 Enterprise
		10.10.10.5     True  True  True pc06.mydomain.com Microsoft Windows Server 2008 R2 Enterprise
		10.10.10.6     True  True  True pc07.mydomain.com Microsoft(R) Windows(R) Server 2003, Standard Edition
		10.10.10.7     True False False
		10.10.10.8    False False False
		10.10.10.9     True  True False pc09.mydomain.com Microsoft Windows Server 2008 R2 Enterprise
		10.10.10.10    True  True False pc10.mydomain.com Microsoft Windows XP Professional

	.EXAMPLE 
		PS C:\> Get-SubNetItems -StartScanIP 10.10.10.2 -Verbose
		VERBOSE: Creating own list class.
		VERBOSE: Start scaning...
		VERBOSE: Starting job (1/20) for 10.10.10.2.
		VERBOSE: Trying get part of data.
		VERBOSE: Trying get last part of data.
		VERBOSE: All jobs is not completed (1/20), please wait... (0)
		VERBOSE: Trying get last part of data.
		VERBOSE: All jobs is not completed (1/20), please wait... (5)
		VERBOSE: Trying get last part of data.
		VERBOSE: All jobs is not completed (1/20), please wait... (10)
		VERBOSE: Trying get last part of data.
		VERBOSE: Geting job 10.10.10.2 result.
		VERBOSE: Removing job 10.10.10.2.
		VERBOSE: Scan finished.


		RunspaceId : d2882105-df8c-4c0a-b92c-0d078bcde752
		Active     : True
		Host       : pc02.mydomain.com
		IP         : 10.10.10.2
		OS_Name    : Microsoft Windows Server 2008 R2 Enterprise
		OS_Ver     : 6.1.7601 Service Pack 1
		WMI        : True
		WinRM      : True
		
	.EXAMPLE 	
		PS C:\> $Result = Get-SubNetItems -StartScanIP 10.10.10.1 -EndScanIP 10.10.10.25 -Ports 80,3389,5900	

		PS C:\> $Result | Select-Object IP, Host, MAC, @{l="Ports";e={[string]::join(", ",($_.Ports | Select-Object @{Label="Ports";Expression={"$($_.Port)-$($_.Status)"}} | Select-Object -ExpandProperty Ports))}} | Format-Table * -AutoSize
		
		IP          Host              MAC               Ports
		--          ----              ---               -----
		10.10.10.1                                      80-False, 3389-False, 5900-False
		10.10.10.2  pc02.mydomain.com 00-15-AD-0C-82-20 80-True, 3389-False, 5900-False
		10.10.10.5  pc05.mydomain.com 00-15-5D-1C-80-25 80-True, 3389-False, 5900-False
		10.10.10.7  pc07.mydomain.com 00-15-4D-0C-81-04 80-True, 3389-True, 5900-False
		10.10.10.9  pc09.mydomain.com 00-15-4A-0C-80-31 80-True, 3389-True, 5900-False
		10.10.10.10 pc10.mydomain.com 00-15-5D-02-1F-1C 80-False, 3389-True, 5900-False

	.NOTES 
		Author: Michal Gajda
		
		ChangeLog:
		v1.3
		-Scan items in subnet for MAC
		-Basic port scan on items in subnet
		-Fixed some small spelling bug
		
		v1.2
		-IP Range Ganerator upgrade
		
		v1.1
		-ProgressBar upgrade
		
		v1.0:
		-Scan subnet for items
		-Scan items in subnet for WMI Access
		-Scan items in subnet for WinRM Access
#>

	[CmdletBinding(
		SupportsShouldProcess=$True,
		ConfirmImpact="Low" 
	)]	
	param(
		[parameter(Mandatory=$true)]
		[System.Net.IPAddress]$StartScanIP,
		[System.Net.IPAddress]$EndScanIP,
		[Int]$MaxJobs = 20,
		[Int[]]$Ports,
		[Switch]$ShowAll,
		[Switch]$ShowInstantly,
		[Int]$SleepTime = 5,
		[Int]$TimeOut = 90
	)

	Begin{}

	Process
	{
		if ($pscmdlet.ShouldProcess("$StartScanIP $EndScanIP" ,"Scan IP range for active machines"))
		{
			if(Get-Job -name *.*.*.*)
			{
				Write-Verbose "Removing old jobs."
				Get-Job -name *.*.*.* | Remove-Job -Force
			}
			
			$ScanIPRange = @()
			if($EndScanIP -ne $null)
			{
				Write-Verbose "Generating IP range list."
				# Many thanks to Dr. Tobias Weltner, MVP PowerShell and Grant Ward for IP range generator
				$StartIP = $StartScanIP -split '\.'
	  			[Array]::Reverse($StartIP)  
	  			$StartIP = ([System.Net.IPAddress]($StartIP -join '.')).Address 
				
				$EndIP = $EndScanIP -split '\.'
	  			[Array]::Reverse($EndIP)  
	  			$EndIP = ([System.Net.IPAddress]($EndIP -join '.')).Address 
				
				For ($x=$StartIP; $x -le $EndIP; $x++) {    
					$IP = [System.Net.IPAddress]$x -split '\.'
					[Array]::Reverse($IP)   
					$ScanIPRange += $IP -join '.' 
				}
			
			}
			else
			{
				$ScanIPRange = $StartScanIP
			}

			Write-Verbose "Creating own list class."
			$Class = @"
			public class SubNetItem {
				public bool Active;
				public string Host;
				public System.Net.IPAddress IP;
				public string MAC;
				public System.Object Ports;
				public string OS_Name;
				public string OS_Ver;
				public bool WMI;
				public bool WinRM;
			}
"@		

			Write-Verbose "Start scaning..."	
			$ScanResult = @()
			$ScanCount = 0
			Write-Progress -Activity "Scan IP Range $StartScanIP $EndScanIP" -Status "Scaning:" -Percentcomplete (0)
			Foreach($IP in $ScanIPRange)
			{
	 			Write-Verbose "Starting job ($((Get-Job -name *.*.*.* | Measure-Object).Count+1)/$MaxJobs) for $IP."
				Start-Job -Name $IP -ArgumentList $IP,$Ports,$Class -ScriptBlock{ 
				
					param
					(
					[System.Net.IPAddress]$IP = $IP,
					[Int[]]$Ports = $Ports,
					$Class = $Class 
					)
					
					Add-Type -TypeDefinition $Class
					
					if(Test-Connection -ComputerName $IP -Quiet)
					{
						#Get Hostname
						Try
						{
							$HostName = [System.Net.Dns]::GetHostbyAddress($IP).HostName
						}
						Catch
						{
							$HostName = $null
						}
						
						#Get WMI Access, OS Name and version via WMI
						Try
						{
							#I don't use Get-WMIObject because it havent TimeOut options. 
							$WMIObj = [WMISearcher]''  
							$WMIObj.options.timeout = '0:0:10' 
							$WMIObj.scope.path = "\\$IP\root\cimv2"  
							$WMIObj.query = "SELECT * FROM Win32_OperatingSystem"  
							$Result = $WMIObj.get()  

							if($Result -ne $null)
							{
								$OS_Name = $Result | Select-Object -ExpandProperty Caption
								$OS_Ver = $Result | Select-Object -ExpandProperty Version
								$OS_CSDVer = $Result | Select-Object -ExpandProperty CSDVersion
								$OS_Ver += " $OS_CSDVer"
								$WMIAccess = $true					
							}
							else
							{
								$WMIAccess = $false	
							}
						}	
						catch
						{
							$WMIAccess = $false					
						}
						
						#Get WinRM Access, OS Name and version via WinRM
						if($HostName)
						{
							$Result = Invoke-Command -ComputerName $HostName -ScriptBlock {systeminfo} -ErrorAction SilentlyContinue 
						}
						else
						{
							$Result = Invoke-Command -ComputerName $IP -ScriptBlock {systeminfo} -ErrorAction SilentlyContinue 
						}
						
						if($Result -ne $null)
						{
							if($OS_Name -eq $null)
							{
								$OS_Name = ($Result[2..3] -split ":\s+")[1]
								$OS_Ver = ($Result[2..3] -split ":\s+")[3]
							}	
							$WinRMAccess = $true
						}
						else
						{
							$WinRMAccess = $false
						}
						
						#Get MAC Address
						Try
						{
							$result= nbtstat -A $IP | select-string "MAC"
							$MAC = [string]([Regex]::Matches($result, "([0-9A-F][0-9A-F]-){5}([0-9A-F][0-9A-F])"))
						}
						Catch
						{
							$MAC = $null
						}
						
						#Get ports status
						$PortsStatus = @()
						ForEach($Port in $Ports)
						{
							Try
							{							
								$TCPClient = new-object Net.Sockets.TcpClient
								$TCPClient.Connect($IP, $Port)
								$TCPClient.Close()
								
								$PortStatus = New-Object PSObject -Property @{            
		        					Port		= $Port
									Status      = $true
								}
								$PortsStatus += $PortStatus
							}	
							Catch
							{
								$PortStatus = New-Object PSObject -Property @{            
		        					Port		= $Port
									Status      = $false
								}	
								$PortsStatus += $PortStatus
							}
						}

						
						$HostObj = New-Object SubNetItem -Property @{            
		        					Active		= $true
									Host        = $HostName
									IP          = $IP 
									MAC         = $MAC
									Ports       = $PortsStatus
		        					OS_Name     = $OS_Name
									OS_Ver      = $OS_Ver               
		        					WMI         = $WMIAccess      
		        					WinRM       = $WinRMAccess      
		        		}
						$HostObj
					}
					else
					{
						$HostObj = New-Object SubNetItem -Property @{            
		        					Active		= $false
									Host        = $null
									IP          = $IP  
									MAC         = $null
									Ports       = $null
		        					OS_Name     = $null
									OS_Ver      = $null               
		        					WMI         = $null      
		        					WinRM       = $null      
		        		}
						$HostObj
					}
				} | Out-Null
				$ScanCount++
				Write-Progress -Activity "Scan IP Range $StartScanIP $EndScanIP" -Status "Scaning:" -Percentcomplete ([int](($ScanCount+$ScanResult.Count)/(($ScanIPRange | Measure-Object).Count) * 50))
				
				do
				{
					Write-Verbose "Trying get part of data."
					Get-Job -State Completed | Foreach {
						Write-Verbose "Geting job $($_.Name) result."
						$JobResult = Receive-Job -Id ($_.Id)

						if($ShowAll)
						{
							if($ShowInstantly)
							{
								if($JobResult.Active -eq $true)
								{
									Write-Host "$($JobResult.IP) is active." -ForegroundColor Green
								}
								else
								{
									Write-Host "$($JobResult.IP) is inactive." -ForegroundColor Red
								}
							}
							
							$ScanResult += $JobResult	
						}
						else
						{
							if($JobResult.Active -eq $true)
							{
								if($ShowInstantly)
								{
									Write-Host "$($JobResult.IP) is active." -ForegroundColor Green
								}
								$ScanResult += $JobResult
							}
						}
						Write-Verbose "Removing job $($_.Name)."
						Remove-Job -Id ($_.Id)
						Write-Progress -Activity "Scan IP Range $StartScanIP $EndScanIP" -Status "Scaning:" -Percentcomplete ([int](($ScanCount+$ScanResult.Count)/(($ScanIPRange | Measure-Object).Count) * 50))
					}
					
					if((Get-Job -name *.*.*.*).Count -eq $MaxJobs)
					{
						Write-Verbose "Jobs are not completed ($((Get-Job -name *.*.*.* | Measure-Object).Count)/$MaxJobs), please wait..."
						Sleep $SleepTime
					}
				}
				while((Get-Job -name *.*.*.*).Count -eq $MaxJobs)
			}
			
			$timeOutCounter = 0
			do
			{
				Write-Verbose "Trying get last part of data."
				Get-Job -State Completed | Foreach {
					Write-Verbose "Geting job $($_.Name) result."
					$JobResult = Receive-Job -Id ($_.Id)

					if($ShowAll)
					{
						if($ShowInstantly)
						{
							if($JobResult.Active -eq $true)
							{
								Write-Host "$($JobResult.IP) is active." -ForegroundColor Green
							}
							else
							{
								Write-Host "$($JobResult.IP) is inactive." -ForegroundColor Red
							}
						}
						
						$ScanResult += $JobResult	
					}
					else
					{
						if($JobResult.Active -eq $true)
						{
							if($ShowInstantly)
							{
								Write-Host "$($JobResult.IP) is active." -ForegroundColor Green
							}
							$ScanResult += $JobResult
						}
					}
					Write-Verbose "Removing job $($_.Name)."
					Remove-Job -Id ($_.Id)
					Write-Progress -Activity "Scan IP Range $StartScanIP $EndScanIP" -Status "Scaning:" -Percentcomplete ([int](($ScanCount+$ScanResult.Count)/(($ScanIPRange | Measure-Object).Count) * 50))
				}
				
				if(Get-Job -name *.*.*.*)
				{
					Write-Verbose "All jobs are not completed ($((Get-Job -name *.*.*.* | Measure-Object).Count)/$MaxJobs), please wait... ($timeOutCounter)"
					Sleep $SleepTime
					$timeOutCounter += $SleepTime				

					if($timeOutCounter -ge $TimeOut)
					{
						Write-Verbose "Time out... $TimeOut. Can't finish some jobs  ($((Get-Job -name *.*.*.* | Measure-Object).Count)/$MaxJobs) try remove it manualy."
						Break
					}
				}
			}
			while(Get-Job -name *.*.*.*)
			
			Write-Verbose "Scan finished."
			Return $ScanResult | Sort-Object {"{0:d3}.{1:d3}.{2:d3}.{3:d3}" -f @([int[]]([string]$_.IP).split('.'))}
		}
	}
	
	End{}
}

function Get-NetworkStatistics {
    <#
    .SYNOPSIS
	    Display current TCP/IP connections for local or remote system

    .FUNCTIONALITY
        Computers

    .DESCRIPTION
	    Display current TCP/IP connections for local or remote system.  Includes the process ID (PID) and process name for each connection.
	    If the port is not yet established, the port number is shown as an asterisk (*).	
	
    .PARAMETER ProcessName
	    Gets connections by the name of the process. The default value is '*'.
	
    .PARAMETER Port
	    The port number of the local computer or remote computer. The default value is '*'.

    .PARAMETER Address
	    Gets connections by the IP address of the connection, local or remote. Wildcard is supported. The default value is '*'.

    .PARAMETER Protocol
	    The name of the protocol (TCP or UDP). The default value is '*' (all)
	
    .PARAMETER State
	    Indicates the state of a TCP connection. The possible states are as follows:
		
	    Closed       - The TCP connection is closed. 
	    Close_Wait   - The local endpoint of the TCP connection is waiting for a connection termination request from the local user. 
	    Closing      - The local endpoint of the TCP connection is waiting for an acknowledgement of the connection termination request sent previously. 
	    Delete_Tcb   - The transmission control buffer (TCB) for the TCP connection is being deleted. 
	    Established  - The TCP handshake is complete. The connection has been established and data can be sent. 
	    Fin_Wait_1   - The local endpoint of the TCP connection is waiting for a connection termination request from the remote endpoint or for an acknowledgement of the connection termination request sent previously. 
	    Fin_Wait_2   - The local endpoint of the TCP connection is waiting for a connection termination request from the remote endpoint. 
	    Last_Ack     - The local endpoint of the TCP connection is waiting for the final acknowledgement of the connection termination request sent previously. 
	    Listen       - The local endpoint of the TCP connection is listening for a connection request from any remote endpoint. 
	    Syn_Received - The local endpoint of the TCP connection has sent and received a connection request and is waiting for an acknowledgment. 
	    Syn_Sent     - The local endpoint of the TCP connection has sent the remote endpoint a segment header with the synchronize (SYN) control bit set and is waiting for a matching connection request. 
	    Time_Wait    - The local endpoint of the TCP connection is waiting for enough time to pass to ensure that the remote endpoint received the acknowledgement of its connection termination request. 
	    Unknown      - The TCP connection state is unknown.
	
	    Values are based on the TcpState Enumeration:
	    http://msdn.microsoft.com/en-us/library/system.net.networkinformation.tcpstate%28VS.85%29.aspx
        
        Cookie Monster - modified these to match netstat output per here:
        http://support.microsoft.com/kb/137984

    .PARAMETER ComputerName
        If defined, run this command on a remote system via WMI.  \\computername\c$\netstat.txt is created on that system and the results returned here

    .PARAMETER ShowHostNames
        If specified, will attempt to resolve local and remote addresses.

    .PARAMETER tempFile
        Temporary file to store results on remote system.  Must be relative to remote system (not a file share).  Default is "C:\netstat.txt"

    .PARAMETER AddressFamily
        Filter by IP Address family: IPv4, IPv6, or the default, * (both).

        If specified, we display any result where both the localaddress and the remoteaddress is in the address family.

    .EXAMPLE
	    Get-NetworkStatistics | Format-Table

    .EXAMPLE
	    Get-NetworkStatistics iexplore -computername k-it-thin-02 -ShowHostNames | Format-Table

    .EXAMPLE
	    Get-NetworkStatistics -ProcessName md* -Protocol tcp

    .EXAMPLE
	    Get-NetworkStatistics -Address 192* -State LISTENING

    .EXAMPLE
	    Get-NetworkStatistics -State LISTENING -Protocol tcp

    .EXAMPLE
        Get-NetworkStatistics -Computername Computer1, Computer2

    .EXAMPLE
        'Computer1', 'Computer2' | Get-NetworkStatistics

    .OUTPUTS
	    System.Management.Automation.PSObject

    .NOTES
	    Author: Shay Levy, code butchered by Cookie Monster
	    Shay's Blog: http://PowerShay.com
        Cookie Monster's Blog: http://ramblingcookiemonster.github.io/

    .LINK
        http://gallery.technet.microsoft.com/scriptcenter/Get-NetworkStatistics-66057d71
    #>	
	[OutputType('System.Management.Automation.PSObject')]
	[CmdletBinding()]
	param(
		
		[Parameter(Position=0)]
		[System.String]$ProcessName='*',
		
		[Parameter(Position=1)]
		[System.String]$Address='*',		
		
		[Parameter(Position=2)]
		$Port='*',

		[Parameter(Position=3,
                   ValueFromPipeline = $True,
                   ValueFromPipelineByPropertyName = $True)]
        [System.String[]]$ComputerName=$env:COMPUTERNAME,

		[ValidateSet('*','tcp','udp')]
		[System.String]$Protocol='*',

		[ValidateSet('*','Closed','Close_Wait','Closing','Delete_Tcb','DeleteTcb','Established','Fin_Wait_1','Fin_Wait_2','Last_Ack','Listening','Syn_Received','Syn_Sent','Time_Wait','Unknown')]
		[System.String]$State='*',

        [switch]$ShowHostnames,
        
        [switch]$ShowProcessNames = $true,	

        [System.String]$TempFile = "C:\netstat.txt",

        [validateset('*','IPv4','IPv6')]
        [string]$AddressFamily = '*'
	)
    
	begin{
        #Define properties
            $properties = 'ComputerName','Protocol','LocalAddress','LocalPort','RemoteAddress','RemotePort','State','ProcessName','PID'

        #store hostnames in array for quick lookup
            $dnsCache = @{}
            
	}
	
	process{

        foreach($Computer in $ComputerName) {

            #Collect processes
            if($ShowProcessNames){
                Try {
                    $processes = Get-Process -ComputerName $Computer -ErrorAction stop | select name, id
                }
                Catch {
                    Write-warning "Could not run Get-Process -computername $Computer.  Verify permissions and connectivity.  Defaulting to no ShowProcessNames"
                    $ShowProcessNames = $false
                }
            }
	    
            #Handle remote systems
                if($Computer -ne $env:COMPUTERNAME){

                    #define command
                        [string]$cmd = "cmd /c c:\windows\system32\netstat.exe -ano >> $tempFile"
            
                    #define remote file path - computername, drive, folder path
                        $remoteTempFile = "\\{0}\{1}`${2}" -f "$Computer", (split-path $tempFile -qualifier).TrimEnd(":"), (Split-Path $tempFile -noqualifier)

                    #delete previous results
                        Try{
                            $null = Invoke-WmiMethod -class Win32_process -name Create -ArgumentList "cmd /c del $tempFile" -ComputerName $Computer -ErrorAction stop
                        }
                        Catch{
                            Write-Warning "Could not invoke create win32_process on $Computer to delete $tempfile"
                        }

                    #run command
                        Try{
                            $processID = (Invoke-WmiMethod -class Win32_process -name Create -ArgumentList $cmd -ComputerName $Computer -ErrorAction stop).processid
                        }
                        Catch{
                            #If we didn't run netstat, break everything off
                            Throw $_
                            Break
                        }

                    #wait for process to complete
                        while (
                            #This while should return true until the process completes
                                $(
                                    try{
                                        get-process -id $processid -computername $Computer -ErrorAction Stop
                                    }
                                    catch{
                                        $FALSE
                                    }
                                )
                        ) {
                            start-sleep -seconds 2 
                        }
            
                    #gather results
                        if(test-path $remoteTempFile){
                    
                            Try {
                                $results = Get-Content $remoteTempFile | Select-String -Pattern '\s+(TCP|UDP)'
                            }
                            Catch {
                                Throw "Could not get content from $remoteTempFile for results"
                                Break
                            }

                            Remove-Item $remoteTempFile -force

                        }
                        else{
                            Throw "'$tempFile' on $Computer converted to '$remoteTempFile'.  This path is not accessible from your system."
                            Break
                        }
                }
                else{
                    #gather results on local PC
                        $results = netstat -ano | Select-String -Pattern '\s+(TCP|UDP)'
                }

            #initialize counter for progress
                $totalCount = $results.count
                $count = 0
    
            #Loop through each line of results    
	            foreach($result in $results) {
            
    	            $item = $result.line.split(' ',[System.StringSplitOptions]::RemoveEmptyEntries)
    
    	            if($item[1] -notmatch '^\[::'){
                    
                        #parse the netstat line for local address and port
    	                    if (($la = $item[1] -as [ipaddress]).AddressFamily -eq 'InterNetworkV6'){
    	                        $localAddress = $la.IPAddressToString
    	                        $localPort = $item[1].split('\]:')[-1]
    	                    }
    	                    else {
    	                        $localAddress = $item[1].split(':')[0]
    	                        $localPort = $item[1].split(':')[-1]
    	                    }
                    
                        #parse the netstat line for remote address and port
    	                    if (($ra = $item[2] -as [ipaddress]).AddressFamily -eq 'InterNetworkV6'){
    	                        $remoteAddress = $ra.IPAddressToString
    	                        $remotePort = $item[2].split('\]:')[-1]
    	                    }
    	                    else {
    	                        $remoteAddress = $item[2].split(':')[0]
    	                        $remotePort = $item[2].split(':')[-1]
    	                    }

                        #Filter IPv4/IPv6 if specified
                            if($AddressFamily -ne "*")
                            {
                                if($AddressFamily -eq 'IPv4' -and $localAddress -match ':' -and $remoteAddress -match ':|\*' )
                                {
                                    #Both are IPv6, or ipv6 and listening, skip
                                    Write-Verbose "Filtered by AddressFamily:`n$result"
                                    continue
                                }
                                elseif($AddressFamily -eq 'IPv6' -and $localAddress -notmatch ':' -and ( $remoteAddress -notmatch ':' -or $remoteAddress -match '*' ) )
                                {
                                    #Both are IPv4, or ipv4 and listening, skip
                                    Write-Verbose "Filtered by AddressFamily:`n$result"
                                    continue
                                }
                            }
    	    		
                        #parse the netstat line for other properties
    	    		        $procId = $item[-1]
    	    		        $proto = $item[0]
    	    		        $status = if($item[0] -eq 'tcp') {$item[3]} else {$null}	

                        #Filter the object
		    		        if($remotePort -notlike $Port -and $localPort -notlike $Port){
                                write-verbose "remote $Remoteport local $localport port $port"
                                Write-Verbose "Filtered by Port:`n$result"
                                continue
		    		        }

		    		        if($remoteAddress -notlike $Address -and $localAddress -notlike $Address){
                                Write-Verbose "Filtered by Address:`n$result"
                                continue
		    		        }
    	    			     
    	    			    if($status -notlike $State){
                                Write-Verbose "Filtered by State:`n$result"
                                continue
		    		        }

    	    			    if($proto -notlike $Protocol){
                                Write-Verbose "Filtered by Protocol:`n$result"
                                continue
		    		        }
                   
                        #Display progress bar prior to getting process name or host name
                            Write-Progress  -Activity "Resolving host and process names"`
                                -Status "Resolving process ID $procId with remote address $remoteAddress and local address $localAddress"`
                                -PercentComplete (( $count / $totalCount ) * 100)
    	    		
                        #If we are running showprocessnames, get the matching name
                            if($ShowProcessNames -or $PSBoundParameters.ContainsKey -eq 'ProcessName'){
                        
                                #handle case where process spun up in the time between running get-process and running netstat
                                if($procName = $processes | Where {$_.id -eq $procId} | select -ExpandProperty name ){ }
                                else {$procName = "Unknown"}

                            }
                            else{$procName = "NA"}

		    		        if($procName -notlike $ProcessName){
                                Write-Verbose "Filtered by ProcessName:`n$result"
                                continue
		    		        }
    	    						
                        #if the showhostnames switch is specified, try to map IP to hostname
                            if($showHostnames){
                                $tmpAddress = $null
                                try{
                                    if($remoteAddress -eq "127.0.0.1" -or $remoteAddress -eq "0.0.0.0"){
                                        $remoteAddress = $Computer
                                    }
                                    elseif($remoteAddress -match "\w"){
                                        
                                        #check with dns cache first
                                            if ($dnsCache.containskey( $remoteAddress)) {
                                                $remoteAddress = $dnsCache[$remoteAddress]
                                                write-verbose "using cached REMOTE '$remoteAddress'"
                                            }
                                            else{
                                                #if address isn't in the cache, resolve it and add it
                                                    $tmpAddress = $remoteAddress
                                                    $remoteAddress = [System.Net.DNS]::GetHostByAddress("$remoteAddress").hostname
                                                    $dnsCache.add($tmpAddress, $remoteAddress)
                                                    write-verbose "using non cached REMOTE '$remoteAddress`t$tmpAddress"
                                            }
                                    }
                                }
                                catch{ }

                                try{

                                    if($localAddress -eq "127.0.0.1" -or $localAddress -eq "0.0.0.0"){
                                        $localAddress = $Computer
                                    }
                                    elseif($localAddress -match "\w"){
                                        #check with dns cache first
                                            if($dnsCache.containskey($localAddress)){
                                                $localAddress = $dnsCache[$localAddress]
                                                write-verbose "using cached LOCAL '$localAddress'"
                                            }
                                            else{
                                                #if address isn't in the cache, resolve it and add it
                                                    $tmpAddress = $localAddress
                                                    $localAddress = [System.Net.DNS]::GetHostByAddress("$localAddress").hostname
                                                    $dnsCache.add($localAddress, $tmpAddress)
                                                    write-verbose "using non cached LOCAL '$localAddress'`t'$tmpAddress'"
                                            }
                                    }
                                }
                                catch{ }
                            }
    
    	    		    #Write the object	
    	    		        New-Object -TypeName PSObject -Property @{
		    		            ComputerName = $Computer
                                PID = $procId
		    		            ProcessName = $procName
		    		            Protocol = $proto
		    		            LocalAddress = $localAddress
		    		            LocalPort = $localPort
		    		            RemoteAddress =$remoteAddress
		    		            RemotePort = $remotePort
		    		            State = $status
		    	            } | Select-Object -Property $properties								

                        #Increment the progress counter
                            $count++
                    }
                }
        }
    }
}

Function set-Switch{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        # WPF File to import
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Market
    )

if(!$market) {
    $market = set-market 
    $selection = ($switches |Out-GridView -Title "Choose a switch" -PassThru).name
}else{
    $selection = ($switches |Where-Object {$_.market -eq $market} |Out-GridView -Title "Choose a switch from $Market" -PassThru).name 
}
return $selection
}

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

function get-market{
    [CmdletBinding()]
    param ( 
        [Parameter(Mandatory=$True,Position=1)] 
        [market]$Market)

    Enum Market{    
        ATL = 1
        ATL1 = 2
        ATL2 = 3
        ATL3 = 4
        CIN = 5
        CIN1 = 6
        CIN2 = 7
        CLT = 8
        CLT1 = 9
        CLT2 = 10
        CLT3 = 11
        CLT4 = 12
        FLL = 13
        FLL1 = 14
        FLL2 = 15
        JAX = 16
        JAX1 = 17
        JAX2 = 18
        LOU = 19
        LOU1 = 20
        LOU2 = 21
        LOU3 = 22
        LOU4 = 23
        LOU5 = 42
        NAS = 24
        NAS1 = 25
        NAS4 = 26
        NAS2 = 27
        NAS3 = 28
        NAS5 = 29
        RAL = 30
        RAL1 = 31
        RAL2 = 32
        RAL3 = 33
        RIC = 34
        RIC1 = 35
        RIC2 = 36
        TPA = 37
        TPA1 = 38
        TPA2 = 39
        TPA3 = 40
        LAB = 41
    }
#Default View
    $defaultProperties = @('Index','mktCode','Fullname','Notes','AccessSwitches','CustCount','vlanCount','VCS','mktCluster','vtp','DNS1','DNS2','Dist1','Dist2','SDDist1','SDDist2','Core1','Core2','VSM','UCSM','dcid','msid','CollectorIP','RC','RouteTargetID')


    #$vlans = import-csv 'S:\Provisioning\Central Provisioning\Automation\Inventory\VLANs.CSV'
    $data = import-csv 'S:\Provisioning\Central Provisioning\Automation\Inventory\MarketList.csv' | where {$_.mktCode -eq $market}
    $access = $switches | where {$_.market -eq $market} | select Name
    $data | add-member -MemberType ScriptProperty -Name "Access" -Value {$switches | where {$_.market -eq $this.mktCode -and $_.type -eq 'Access'} | select Name, EM7_ID, Po}
    $data | add-member -MemberType ScriptProperty -Name "SD" -Value {$switches | where {$_.market -eq $this.mktCode -and $_.type -eq 'SD'} | select Name, EM7_ID}
    $data | add-member -MemberType ScriptProperty -name "AccessSwitches" -Value {[int]($this.access).count}
    $data | Add-Member -MemberType ScriptProperty -force -name "VLANs" -Value { $vlans | Where-Object {$_.market -eq $this.mktcode} | select vlan, custid, desc, market }
    $data | add-member -MemberType ScriptProperty -force -name "Customers" -value { $vlans | Where-Object {$_.market -eq $this.mktcode -and $_.custid -notlike "*PEAK*"} |select -Unique CustID}
    $data | add-member -MemberType ScriptProperty -Force -name "VPL" -value {$vlans | Where-Object {$_.market -eq $this.mktcode -and $_.desc -like "VPL*"} | select vlan, custid, desc, market}
#Free Access Ports
    $data | add-member -MemberType ScriptMethod -name "GetFreeAccessPorts" -Value {
        $a = get-FreePorts -market $this.mktcode -SwitchType Access
        $Fports = $a |select hostname, name
        $this | ForEach-Object {$_ | add-member -name 'FreePorts' -value $fports -MemberType NoteProperty}
        $this.freeports
        $this | ForEach-Object {$_ |add-member -MemberType NoteProperty -name "UnusedPorts" -value $a.count        }
        $mktcode = $this.mktcode 
        $mkt | Where-Object {$_.market -eq $mktcode} |ForEach-Object {$_ | Add-Member -MemberType NoteProperty -name "FreePortCount" -value $a.count -force}
        $defaultProperties = @('Index','mktCode','Fullname','Notes','AccessSwitches','CustCount','vlanCount','VCS','mktCluster','vtp','DNS1','DNS2','Dist1','Dist2','SDDist1','SDDist2','Core1','Core2','VSM','UCSM','dcid','msid','CollectorIP','RC','RouteTargetID','UnusedPorts')
        $sd = get-FreePorts -market $this.mktcode -SwitchType SD
        $this | add-member -MemberType NoteProperty -name 'SDFreePorts' -value $sd
     }


#get Customers
    $data | add-member -MemberType ScriptMethod -Name GetCustomers -Value {
        $all = Get-netMktCustomers -market $this.mktcode -ErrorAction SilentlyContinue
        $this | add-member -MemberType NoteProperty -Name 'Customers' -Value $all
        $this | add-member -MemberType NoteProperty -name 'CustomerCount' -Value $all.count
    }
#get VLANs
    $data | Add-Member -MemberType ScriptMethod -name GetVLANs -value {
        $vlans = get-sVLAN -Hostname $this.dist1, $this.sddist1
        $this | add-member -MemberType NoteProperty -Force -name 'VLANs' -Value { $vlans |select VLAN, CustID, Desc, market} 
        write-host "Added VLAN list to Market member.
        Total: $($vlans.count) VLANs" -ForegroundColor Green
        #return $vlans 
    }
    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultProperties)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    $data | Add-Member MemberSet PSStandardMembers $PSStandardMembers
    $data | add-member ScriptProperty -Force -Name CustCount -Value {[int]($this.customers | select -Unique CustID).count}
    $data | add-member ScriptProperty -Force -Name vlanCount -value {[int]($this.vlans).count}
    $data | add-member -MemberType ScriptMethod -Name vlan -Value {
        param([string]$parameter01 = $(throw "Must supply a searchstring for VLAN description."))
        $p = "*$($parameter01)*"
        $a = $this.vlans | where {$_.desc -like $p}
        return $a
    }
        $data | add-member -MemberType ScriptMethod -Name CustID -Value {
        param([string]$parameter01 = $(throw "Must supply a searchstring for Customer ID."))
        $p = "*$($parameter01)*"
        $a = $this.customers | where {$_.custid -like $p}
        return $a
    }

return $data
}

function Start-Putty{
        [CmdletBinding()]
        [Alias()]
        param(
        [Parameter(Mandatory=$True,Position=1)] 
        [string]$Hostname, 
        [Parameter(Mandatory=$False,Position=2)] 
        [String]$user,
        [Parameter(Mandatory=$False,Position=2)]
        [string]$password 
        )

    $user = $env:USERNAME
    $pass = if(!$password)  {
        $creds = get-credential -Message 'Enter Password for switch' -UserName $user
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($creds.password)
        $pass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        putty.exe -ssh -load $hostname -l $user -pw $Pass 
    }
    else{
        putty.exe -ssh -load $hostname -l $user -pw $password    
    }
}

function Start-cPutty{
        [CmdletBinding()]
        [Alias()]
        param(
        [Parameter(Mandatory=$True,Position=1,ValueFromPipelineByPropertyName,ValueFromPipeline)] 
        [string]$Hostname
        )


    if(!$creds){$global:creds = Get-Credential -UserName $env:USERNAME -Message 'Enter Password'}
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($creds.password)
    $pass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)


    $user = $env:USERNAME
            $cString = "$($user)@$($hostname).peak10.net"
        putty.exe -new_console:s -ssh $cString -pw $Pass 
}

function invoke-markets {
    
    $mkt = $switches | Where-Object {$_.market -ne 'LOU5'}| select -unique market
    $vars = get-variable | select name
    foreach ($m in $mkt) {
        if($m.market -notin $vars.name){
            $ErrorActionPreference = 'SilentlyContinue'
            $a = get-market -Market $m.market
            new-variable -Name $($m.market) -value $a -Scope Global
            $ErrorActionPreference = 'Continue'
        }else{$m}
    }
    $mkt | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -Name ASwitchCount -Value (Get-Variable $($_.market)).Value.accessswitches -force}
    $mkt | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -Name Dist -Value (Get-Variable $($_.market)).Value.dist1 -force}
    $mkt | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -Name SDDist -Value (Get-Variable $($_.market)).Value.sddist1 -force}
    $mkt | Add-Member -MemberType ScriptMethod -name UpdateCounts -Value { $mkt | Add-member -MemberType NoteProperty -name 'FreePortCount' -value Get-Variable $($_.market).Value.freeports -force}
    
}


function New-HTMLReport{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        [Parameter(Mandatory=$True,ValueFromPipelineByPropertyName=$true,Position=1)]
        $data,
        [Parameter(Mandatory=$false,Position=2)]
        $Head,
        [Parameter(Mandatory=$true,Position=3)]
        $Path

    )
    $fragments = @()    
    $fragments+= "<H1>$($head)</H1>"
    [xml]$html = $data | convertto-html -Fragment
 
    for ($i=1;$i -le $html.table.tr.count-1;$i++) {
      if ($html.table.tr[$i].td[10] -eq 'SVI') {
        $class = $html.CreateAttribute("class")
        $class.value = 'alert'
        $html.table.tr[$i].childnodes[3].attributes.append($class) | out-null
      }
    }

    $fragments+= $html.InnerXml

    $fragments+= "<p class='footer'>$(get-date)</p>"
    $convertParams = @{ 
        head = @"
    <Title>Customer Networks</Title>
<style>
body { background-color:#E5E4E2;
        font-family:Monospace;
        font-size:10pt; }
td, th { border:0px solid black; 
            border-collapse:collapse;
            white-space:pre; }
th { color:white;
        background-color:black; }
table, tr, td, th { padding: 2px; margin: 0px ;white-space:pre; }
tr:nth-child(odd) {background-color: lightgray}
table { width:95%;margin-left:5px; margin-bottom:20px;}
h2 {
    font-family:Tahoma;
    color:#6D7B8D;
}
.alert {
    color: red; 
    }
.footer 
{ color:green; 
    margin-left:10px; 
    font-family:Tahoma;
    font-size:8pt;
    font-style:italic;
}
</style>
"@
    body = $fragments
    }
    convertto-html @convertParams | out-file $path -Append
}


function ToArray
{
  begin
  {
    $output = @();
  }
  process
  {
    $output += $_;
  }
  end
  {
    return ,$output;
  }
}

function select-market{

$mkt_selected = StackPanel -ControlName 'Market' {
    new-label -VisualStyle 'mediumText' "Market"
    New-ComboBox -IsEditable:$false -SelectedIndex 0 -Name Market @($mkt.market)
    New-Button "Select Market" -On_Click {            
        Get-ParentControl |            
            Set-UIValue -passThru |             
            Close-Control            
    }            
} -show
return $mkt_selected.market.tostring()

}

function select-switch{
    [CmdletBinding()]
    [Alias()]
    [OutputType([string])]
    Param
    (
        [Parameter(Mandatory=$True,ValueFromPipelineByPropertyName=$true,Position=1)]
        $Market
    )
    $switch_selected = StackPanel -ControlName 'Switch' {
        new-label -VisualStyle 'mediumText' "switch"
        New-ComboBox -IsEditable:$false -SelectedIndex 0 -Name Switch @(($switches | where {$_.market -eq $market}).name)
        New-Button "Get Customer" -On_Click {            
            Get-ParentControl |            
                Set-UIValue -passThru |             
                Close-Control            
        }            
    } -show
    return $switch_selected
}
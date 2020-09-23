Attribute VB_Name = "Module1"
Option Explicit
Public Clones As Integer
Public Running As Boolean
Public Timeout As Integer
Public Range As Long

Public Const PING_TIMEOUT = 1000
Public Const WS_VERSION_REQD = &H101
Public Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long

Declare Function gethostname Lib "WSOCK32.DLL" (ByVal hostname$, HostLen&) As Long
Declare Function gethostbyname& Lib "WSOCK32.DLL" (ByVal hostname$)
Declare Function gethostbyaddr& Lib "WSOCK32.DLL" (ByVal adr$, length%, cType%)
Declare Function inet_addr& Lib "WSOCK32.DLL" (ByVal cp$)

Public Declare Function WSAStartup Lib "WSOCK32.DLL" _
   (ByVal wVersionRequired As Long, _
    lpWSADATA As WSADATA) As Long
    
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function IsDestinationReachable Lib "SENSAPI.DLL" Alias "IsDestinationReachableA" (ByVal lpszDestination As String, ByRef lpQOCInfo As QOCINFO) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Declare Function IcmpCreateFile Lib "icmp.dll" () As Long

Public Declare Function IcmpCloseHandle Lib "icmp.dll" _
   (ByVal IcmpHandle As Long) As Long
   
Public Declare Function IcmpSendEcho Lib "icmp.dll" _
   (ByVal IcmpHandle As Long, _
    ByVal DestinationAddress As Long, _
    ByVal RequestData As String, _
    ByVal RequestSize As Integer, _
    ByVal RequestOptions As Long, _
    ReplyBuffer As ICMP_ECHO_REPLY, _
    ByVal ReplySize As Long, _
    ByVal Timeout As Long) As Long
    

Public Const GWL_EXSTYLE = (-20)
Public Const LWA_ALPHA = &H2
Public Const WS_EX_LAYERED = &H80000

Private Type QOCINFO
  dwSize As Long
  dwFlags As Long
  dwInSpeed As Long 'in bytes/second
  dwOutSpeed As Long 'in bytes/second
End Type

Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 255) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Type WebName
    Byte1(1 To 32) As Byte
    name As String * 64
    Byte2(1 To 32) As Byte
End Type

Public Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type

Public Type ICMP_ECHO_REPLY
    Address         As Long
    status          As Long
    RoundTripTime   As Long
    DataSize        As Integer
    Reserved        As Integer
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type


Type Hostent
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Type Inet_address
    Byte4 As String * 1
    Byte3 As String * 1
    Byte2 As String * 1
    Byte1 As String * 1
End Type


Function AddressStringToLong(ByVal tmp As String) As Long

   Dim i As Integer
   Dim parts(1 To 4) As String
   
   i = 0
   
  'we have to extract each part of the
  '123.456.789.123 string, delimited by
  'a period
   While InStr(tmp, ".") > 0
      i = i + 1
      parts(i) = Mid(tmp, 1, InStr(tmp, ".") - 1)
      tmp = Mid(tmp, InStr(tmp, ".") + 1)
   Wend
   
   i = i + 1
   parts(i) = tmp
   
   If i <> 4 Then
      AddressStringToLong = 0
      Exit Function
   End If
   
  'build the long value out of the
  'hex of the extracted strings
   AddressStringToLong = Val("&H" & Right("00" & Hex(parts(4)), 2) & _
                         Right("00" & Hex(parts(3)), 2) & _
                         Right("00" & Hex(parts(2)), 2) & _
                         Right("00" & Hex(parts(1)), 2))
   
End Function
Public Function Transparent(ByVal hWnd As Long, Perc As Integer) As Long
    Dim Msg As Long
    On Error Resume Next
    If Perc < 0 Or Perc > 255 Then
      Transparent = 1
    Else
      Msg = GetWindowLong(hWnd, GWL_EXSTYLE)
      Msg = Msg Or WS_EX_LAYERED
      SetWindowLong hWnd, GWL_EXSTYLE, Msg
      SetLayeredWindowAttributes hWnd, 0, Perc, LWA_ALPHA
      Transparent = 0
    End If
    If err Then
      Transparent = 2
    End If
End Function
Public Function Ping(szAddress As String, ECHO As ICMP_ECHO_REPLY) As Long
Dim hPort As Long
Dim dwAddress As Long
Dim sDataToSend As String
Dim iOpt As Long
Dim WSAD As WSADATA

   sDataToSend = String(32, 0&)
   dwAddress = AddressStringToLong(szAddress)
   
   
   Call WSAStartup(WS_VERSION_REQD, WSAD)
   hPort = IcmpCreateFile()
   
   If IcmpSendEcho(hPort, _
                   dwAddress, _
                   sDataToSend, _
                   Len(sDataToSend), _
                   0, _
                   ECHO, _
                   Len(ECHO), _
                   PING_TIMEOUT) Then
   
        'the ping succeeded,
        '.Status will be 0
        '.RoundTripTime is the time in ms for
        '               the ping to complete,
        '.Data is the data returned (NULL terminated)
        '.Address is the Ip address that actually replied
        '.DataSize is the size of the string in .Data
         Ping = ECHO.RoundTripTime
   Else: Ping = ECHO.status * -1
   End If
                       
   Call IcmpCloseHandle(hPort)
   WSACleanup


End Function


Public Function vbgetHostName(IP As String) As String
Dim pointer As Long
Dim hostEntity As Hostent
Dim adr As Long
Dim Address As Inet_address
Dim host As WebName
Dim chradr As String
Dim pos As Integer
On Error GoTo err

adr = inet_addr(IP)
CopyMemory Address, adr, Len(Address)
chradr = Address.Byte4 & Address.Byte3 & Address.Byte2 & Address.Byte1
pointer = gethostbyaddr(chradr, Len(chradr), 2)

If pointer = 0 Then Exit Function
CopyMemory hostEntity, pointer, Len(hostEntity)
CopyMemory host, ByVal hostEntity.h_name, Len(host)


pos = InStr(1, host.name, Chr(0), vbBinaryCompare)
host.name = Mid(host.name, 1, pos)

'Debug.Print host.name


vbgetHostName = RTrim(host.name)
Exit Function

err: MsgBox err.Description, vbCritical, "error"

End Function



Attribute VB_Name = "modWinsock"
'
' Generic Client Server class objects functions module 1.1
'
' Copyright 2005, E.V.I.C.T. B.V.
' Website: http:\\www.evict.nl
' Support: mailto:evict@vermeer.nl
'
'Purpose
' For limiting memory load as much as possible functions are moved to this module.
'
'License:
' GPL - The GNU General Public License
' Permits anyone the right to use and modify the software without limitations
' as long as proper credits are given and the original and modified source code
' are included. Requires that the final product, software derivate from the
' original source or any software utilizing a GPL component, such as this,
' is also licensed under the GPL license.
' For more information see http://www.gnu.org/licenses/gpl.txt
'
'License adition:
' You are permitted to use the software in a non-commercial context free of
' charge as long as proper credits are given and the original unmodified source
' code is included.
' For more information see http://www.evict.nl/licenses.html
'
'License exeption:
' If you would like to obtain a commercial license then please contact E.V.I.C.T. B.V.
' For more information see http://www.evict.nl/licenses.html
'
'Terms:
' This software is provided "as is", without warranty of any kind, express or
' implied, including  but not limited to the warranties of merchantability,
' fitness for a particular purpose and noninfringement. In no event shall the
' authors or copyright holders be liable for any claim, damages or other
' liability, whether in an action of contract, tort or otherwise, arising
' from, out of or in connection with the software or the use or other
' dealings in the software.
'
'History:
' 2002 : Created and added to the sharware library siteskinner
' jan 2005 : Changed the licensing from shareware to opensource
' feb 2005 : Added the SetSockLinger and getascip for the Improved connection method

Option Explicit

'Winsock Initialization and termination
Public Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSAData) As Long
Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long

'Server side Winsock API functions
Public Declare Function WSABind Lib "ws2_32.dll" Alias "bind" (ByVal s As Long, ByRef Name As SOCKADDR_IN, ByRef namelen As Long) As Long
Public Declare Function WSAListen Lib "ws2_32.dll" Alias "listen" (ByVal s As Long, ByVal backlog As Long) As Long
Public Declare Function WSAAccept Lib "ws2_32.dll" Alias "accept" (ByVal s As Long, ByRef addr As SOCKADDR_IN, ByRef addrlen As Long) As Long

'String functions
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long

'Socket Functions
Public Declare Function Socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
Public Declare Function WSAConnect Lib "ws2_32.dll" Alias "connect" (ByVal s As Long, ByRef Name As SOCKADDR_IN, ByVal namelen As Long) As Long
Public Declare Function WSACloseSocket Lib "ws2_32.dll" Alias "closesocket" (ByVal s As Long) As Long
Public Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Public Declare Function WSAIsBlocking Lib "wsock32.dll" () As Long
Public Declare Function WSACancelBlockingCall Lib "wsock32.dll" () As Long
Public Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Public Declare Function WSAAsyncGetHostByName Lib "ws2_32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal strHostName As String, buf As Any, ByVal buflen As Long) As Long
Public Declare Function WSARecv Lib "ws2_32.dll" Alias "recv" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Public Declare Function WSASend Lib "ws2_32.dll" Alias "send" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Public Declare Function WSAsetsockopt Lib "wsock32.dll" Alias "setsockopt" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Public Declare Function WSAgetsockopt Lib "wsock32.dll" Alias "getsockopt" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long

'Network byte ordering functions
Public Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer
Public Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
Public Declare Function ntohl Lib "ws2_32.dll" (ByVal netlong As Long) As Long
Public Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Integer) As Integer

'End point information
Public Declare Function getsockname Lib "ws2_32.dll" (ByVal s As Long, ByRef Name As SOCKADDR_IN, ByRef namelen As Long) As Long
Public Declare Function getpeername Lib "ws2_32.dll" (ByVal s As Long, ByRef Name As SOCKADDR_IN, ByRef namelen As Long) As Long

'Hostname resolving functions
Public Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long

'ICMP functions
Public Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Public Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Public Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Long, ByVal RequestOptions As Long, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Long

'Winsock API functions for resolving hostnames and IP's
Public Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal inn As Long) As Long
Public Declare Function gethostbyaddr Lib "wsock32.dll" (haddr As Long, ByVal hnlen As Long, ByVal addrtype As Long) As Long
Public Declare Function gethostname Lib "wsock32.dll" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Public Declare Function gethostbyname Lib "wsock32.dll" (ByVal szHost As String) As Long

'Memory copy and move functions
Public Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Any, ByVal cbCopy As Long)

'Window creation and destruction functions
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

'Messaging functions
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

'Memory allocation functions
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

'..
Private Const WSADESCRIPTION_LEN = 257
Private Const WSASYS_STATUS_LEN = 129

'Maximum queue length specifiable by listen.
Public Const SOMAXCONN = &H7FFFFFFF

'Windows Socket types
Public Const SOCK_STREAM = 1     'Stream socket

'Address family
Public Const AF_INET = 2          'Internetwork: UDP, TCP, etc.

'Socket Protocol
Public Const IPPROTO_TCP = 6     'tcp

'Data type conversion constants
Public Const OFFSET_4 = 4294967296#
Public Const MAXINT_4 = 2147483647
Public Const OFFSET_2 = 65536
Public Const MAXINT_2 = 32767

'Fixed memory flag for GlobalAlloc
Public Const GMEM_FIXED = &H0

'Winsock error offset
Public Const WSABASEERR = 10000

' Other constants
Public Const ERROR_SUCCESS              As Long = 0
Public Const MIN_SOCKETS_REQD           As Long = 1
Public Const WS_VERSION_REQD            As Long = &H101
Public Const WS_VERSION_MAJOR           As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR           As Long = WS_VERSION_REQD And &HFF&
Public Const DATA_SIZE = 32
Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128
Public Const SOL_SOCKET = &HFFFF&
Public Const SO_LINGER = &H80&
Public Const hostent_size = 16
Public Const sockaddr_size = 16

Public Enum IP_STATUS
    IP_STATUS_BASE = 11000
    IP_SUCCESS = 0
    IP_BUF_TOO_SMALL = (11000 + 1)
    IP_DEST_NET_UNREACHABLE = (11000 + 2)
    IP_DEST_HOST_UNREACHABLE = (11000 + 3)
    IP_DEST_PROT_UNREACHABLE = (11000 + 4)
    IP_DEST_PORT_UNREACHABLE = (11000 + 5)
    IP_NO_RESOURCES = (11000 + 6)
    IP_BAD_OPTION = (11000 + 7)
    IP_HW_ERROR = (11000 + 8)
    IP_PACKET_TOO_BIG = (11000 + 9)
    IP_REQ_TIMED_OUT = (11000 + 10)
    IP_BAD_REQ = (11000 + 11)
    IP_BAD_ROUTE = (11000 + 12)
    IP_TTL_EXPIRED_TRANSIT = (11000 + 13)
    IP_TTL_EXPIRED_REASSEM = (11000 + 14)
    IP_PARAM_PROBLEM = (11000 + 15)
    IP_SOURCE_QUENCH = (11000 + 16)
    IP_OPTION_TOO_BIG = (11000 + 17)
    IP_BAD_DESTINATION = (11000 + 18)
    IP_ADDR_DELETED = (11000 + 19)
    IP_SPEC_MTU_CHANGE = (11000 + 20)
    IP_MTU_CHANGE = (11000 + 21)
    IP_UNLOAD = (11000 + 22)
    IP_ADDR_ADDED = (11000 + 23)
    IP_GENERAL_FAILURE = (11000 + 50)
    MAX_IP_STATUS = 11000 + 50
    IP_PENDING = (11000 + 255)
    PING_TIMEOUT = 255
End Enum

'Winsock messages that will go to the window handler
Public Enum WSAMessage
    FD_READ = &H1&      'Data is ready to be read from the buffer
    FD_WRITE = &H2&
    FD_CONNECT = &H10&  'Connection esatblished
    FD_CLOSE = &H20&    'Connection closed
    FD_ACCEPT = &H8&    'Connection request pending
End Enum

Public Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type

Public Type ICMP_ECHO_REPLY
    Address         As Long
    Status          As Long
    RoundTripTime   As Long
    DataSize        As Long
    ' Reserved        As Integer
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type

'Winsock Data structure
Public Type WSAData
    wVersion       As Integer                       'Version
    wHighVersion   As Integer                       'High Version
    szDescription  As String * WSADESCRIPTION_LEN   'Description
    szSystemStatus As String * WSASYS_STATUS_LEN    'Status of system
    iMaxSockets    As Integer                       'Maximum number of sockets allowed
    iMaxUdpDg      As Integer                       'Maximum UDP datagrams
    lpVendorInfo   As Long                          'Vendor Info
End Type

'HostEnt Structure
Public Type HOSTENT
    hName     As Long       'Host Name
    hAliases  As Long       'Alias
    hAddrType As Integer    'Address Type
    hLength   As Integer    'Length
    hAddrList As Long       'Address List
End Type

'Socket Address structure
Public Type SOCKADDR_IN
    sin_family       As Integer 'Address familly
    sin_port         As Integer 'Port
    sin_addr         As Long    'Long address
    sin_zero(1 To 8) As Byte
End Type

'End Point of connection information
Public Enum IPEndPointFields
    LOCAL_HOST          'Local hostname
    LOCAL_HOST_IP       'Local IP
    LOCAL_PORT          'Local port
    REMOTE_HOST         'Remote hostname
    REMOTE_HOST_IP      'Remote IP
    REMOTE_PORT         'Remote port
End Enum

'Basic Winsock error results.
Public Enum WSABaseErrors
    INADDR_NONE = &HFFFF
    SOCKET_ERROR = -1
    INVALID_SOCKET = -1
End Enum

Type LingerType
    l_onoff As Integer
    l_linger As Integer
End Type

'Winsock error constants
Public Enum WSAErrorConstants
    'Windows Sockets definitions of regular Microsoft C error constants
    WSAEINTR = (WSABASEERR + 4)
    WSAEBADF = (WSABASEERR + 9)
    WSAEACCES = (WSABASEERR + 13)
    WSAEFAULT = (WSABASEERR + 14)
    WSAEINVAL = (WSABASEERR + 22)
    WSAEMFILE = (WSABASEERR + 24)
    'Windows Sockets definitions of regular Berkeley error constants
    WSAEWOULDBLOCK = (WSABASEERR + 35)
    WSAEINPROGRESS = (WSABASEERR + 36)
    WSAEALREADY = (WSABASEERR + 37)
    WSAENOTSOCK = (WSABASEERR + 38)
    WSAEDESTADDRREQ = (WSABASEERR + 39)
    WSAEMSGSIZE = (WSABASEERR + 40)
    WSAEPROTOTYPE = (WSABASEERR + 41)
    WSAENOPROTOOPT = (WSABASEERR + 42)
    WSAEPROTONOSUPPORT = (WSABASEERR + 43)
    WSAESOCKTNOSUPPORT = (WSABASEERR + 44)
    WSAEOPNOTSUPP = (WSABASEERR + 45)
    WSAEPFNOSUPPORT = (WSABASEERR + 46)
    WSAEAFNOSUPPORT = (WSABASEERR + 47)
    WSAEADDRINUSE = (WSABASEERR + 48)
    WSAEADDRNOTAVAIL = (WSABASEERR + 49)
    WSAENETDOWN = (WSABASEERR + 50)
    WSAENETUNREACH = (WSABASEERR + 51)
    WSAENETRESET = (WSABASEERR + 52)
    WSAECONNABORTED = (WSABASEERR + 53)
    WSAECONNRESET = (WSABASEERR + 54)
    WSAENOBUFS = (WSABASEERR + 55)
    WSAEISCONN = (WSABASEERR + 56)
    WSAENOTCONN = (WSABASEERR + 57)
    WSAESHUTDOWN = (WSABASEERR + 58)
    WSAETOOMANYREFS = (WSABASEERR + 59)
    WSAETIMEDOUT = (WSABASEERR + 60)
    WSAECONNREFUSED = (WSABASEERR + 61)
    WSAELOOP = (WSABASEERR + 62)
    WSAENAMETOOLONG = (WSABASEERR + 63)
    WSAEHOSTDOWN = (WSABASEERR + 64)
    WSAEHOSTUNREACH = (WSABASEERR + 65)
    WSAENOTEMPTY = (WSABASEERR + 66)
    WSAEPROCLIM = (WSABASEERR + 67)
    WSAEUSERS = (WSABASEERR + 68)
    WSAEDQUOT = (WSABASEERR + 69)
    WSAESTALE = (WSABASEERR + 70)
    WSAEREMOTE = (WSABASEERR + 71)
    'Extended Windows Sockets error constant definitions
    WSASYSNOTREADY = (WSABASEERR + 91)
    WSAVERNOTSUPPORTED = (WSABASEERR + 92)
    WSANOTINITIALISED = (WSABASEERR + 93)
    WSAEDISCON = (WSABASEERR + 101)
    WSAENOMORE = (WSABASEERR + 102)
    WSAECANCELLED = (WSABASEERR + 103)
    WSAEINVALIDPROCTABLE = (WSABASEERR + 104)
    WSAEINVALIDPROVIDER = (WSABASEERR + 105)
    WSAEPROVIDERFAILEDINIT = (WSABASEERR + 106)
    WSASYSCALLFAILURE = (WSABASEERR + 107)
    WSASERVICE_NOT_FOUND = (WSABASEERR + 108)
    WSATYPE_NOT_FOUND = (WSABASEERR + 109)
    WSA_E_NO_MORE = (WSABASEERR + 110)
    WSA_E_CANCELLED = (WSABASEERR + 111)
    WSAEREFUSED = (WSABASEERR + 112)
    WSAHOST_NOT_FOUND = 11001
    WSATRY_AGAIN = 11002
    WSANO_RECOVERY = 11003
    WSANO_DATA = 11004
    FD_SETSIZE = 64
End Enum

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Convert an unsigned long to an integer.

Public Function UnsignedToInteger(Value As Long) As Integer

1     On Error GoTo ErrorHandler

2     If Value < 0 Or Value >= OFFSET_2 Then Error 6  'Overflow

3     If Value <= MAXINT_2 Then
4         UnsignedToInteger = Value
5     Else
6         UnsignedToInteger = Value - OFFSET_2
7     End If

8 Exit Function

9 ErrorHandler:
10     Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in UnsignedToInteger on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Function

' Convert an integer to an unsigned long.

Public Function IntegerToUnsigned(Value As Integer) As Long

11     On Error GoTo ErrorHandler

12     If Value < 0 Then
13         IntegerToUnsigned = Value + OFFSET_2
14     Else
15         IntegerToUnsigned = Value
16     End If

17 Exit Function

18 ErrorHandler:
19     Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in IntegerToUnsigned on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Function

' Create a string from a pointer

Public Function StringFromPointer(ByVal lngPointer As Long) As String

20     On Error GoTo ErrorHandler

21 Dim strTemp As String
22 Dim lRetVal As Long

23     strTemp = String$(lstrlen(ByVal lngPointer), 0)    'prepare the strTemp buffer
24     lRetVal = lstrcpy(ByVal strTemp, ByVal lngPointer) 'copy the string into the strTemp buffer
25     If lRetVal Then StringFromPointer = strTemp        'return the string

26 Exit Function

27 ErrorHandler:
28     Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in StringFromPointer on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Function

' Return the Hi Word of a long value.

Public Function HiWord(lngValue As Long) As Long

29     On Error GoTo ErrorHandler

30     If (lngValue And &H80000000) = &H80000000 Then
31         HiWord = ((lngValue And &H7FFF0000) \ &H10000) Or &H8000&
32     Else
33         HiWord = (lngValue And &HFFFF0000) \ &H10000
34     End If

35 Exit Function

36 ErrorHandler:
37     Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in HiWord on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Function

' Get the received data from the socket and return it in the calling string.
' This will use a byte array to Unicode conversion.

Public Function mRecv(ByVal lngSocket As Long, ByRef strBuffer As String) As Long

38     On Error GoTo ErrorHandler

39 Const MAX_BUFFER_LENGTH As Long = 8192 'Normal= 8192  'MAX = 65536

40 Dim arrBuffer(1 To MAX_BUFFER_LENGTH)   As Byte
41 Dim lngBytesReceived                    As Long
42 Dim strTempBuffer                       As String

    'Call the recv Winsock API function in order to read data from the buffer
43     lngBytesReceived = WSARecv(lngSocket, arrBuffer(1), MAX_BUFFER_LENGTH, 0&)

44     If lngBytesReceived > 0 Then
        'If we have received some data, convert it to the Unicode
        'string that is suitable for the Visual Basic String data type
45         strTempBuffer = StrConv(arrBuffer, vbUnicode)

        'Remove unused bytes
46         strBuffer = Left$(strTempBuffer, lngBytesReceived)
47     End If

48     mRecv = lngBytesReceived

49 Exit Function

50 ErrorHandler:
51     Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in mRecv on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Function

' Send data to the specified socket.
' This will use a Unicode to byte array conversion

Public Function mSend(ByVal lngSocket As Long, strData As String) As Long

52 Dim sockerror As Long

53     On Error GoTo ErrorHandler

54 Dim arrBuffer()     As Byte

    'Convert the data string to a byte array
55     arrBuffer() = StrConv(strData, vbFromUnicode)
    'Call the send Winsock API function in order to send data

56     DoEvents
57     mSend = WSASend(lngSocket, arrBuffer(0), Len(strData), 0&)
58     DoEvents
59     If mSend = SOCKET_ERROR Then
60         sockerror = WSAGetLastError()
61         If sockerror > 0 Then Err.Raise sockerror + 8000, "WSASend", "Winsock error " & sockerror & " wile sending byte data. " & vbCrLf & GetErrorDescription(sockerror)
62     End If

63 Exit Function

64 ErrorHandler:
65     Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in mSend on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Function

' Get the received data from the socket and return it in the calling string.
' This will be done without any unicode conversion.

Public Function mRecvByte(ByVal lngSocket As Long, ByRef byteData() As Byte) As Long

66     On Error GoTo ErrorHandler
67 Dim MAX_BUFFER_LENGTH As Long  ' 2 'Normal= 8192  'MAX = 65536

68     MAX_BUFFER_LENGTH = UBound(byteData())
69     If MAX_BUFFER_LENGTH > 65536 Then MAX_BUFFER_LENGTH = 65536
    'Call the recv Winsock API function in order to read data from the buffer
70     mRecvByte = WSARecv(lngSocket, byteData(LBound(byteData)), MAX_BUFFER_LENGTH, 0&)

71 Exit Function

72 ErrorHandler:
73     Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in mRecvByte on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Function

' Send data to the specified socket.
' This will be done without any unicode conversion.

Public Function mSendByte(ByVal lngSocket As Long, ByRef byteData() As Byte) As Long

74 Dim sockerror As Long

75     On Error GoTo ErrorHandler

    'Call the send Winsock API function in order to send data
76     DoEvents
77     mSendByte = WSASend(lngSocket, byteData(LBound(byteData)), UBound(byteData) - LBound(byteData) + 1, 0&)
78     DoEvents
79     If mSendByte = SOCKET_ERROR Then
80         sockerror = WSAGetLastError()
81         If sockerror > 0 Then Err.Raise sockerror + 9000, "WSASend", "Winsock error " & sockerror & " wile sending byte data. " & vbCrLf & GetErrorDescription(sockerror)
82     End If

83 Exit Function

84 ErrorHandler:
85     Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in mSendByte on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Function

' Get the IP adress of an endpoint (client or server).

Public Function GetIPEndPointField(ByVal lngSocket As Long, ByVal EndpointField As IPEndPointFields) As Variant

86     On Error GoTo ErrorHandler

87 Dim udtSocketAddress    As SOCKADDR_IN
88 Dim lngReturnValue      As Long
89 Dim lngPtrToAddress     As Long
90 Dim strIPAddress        As String
91 Dim lngAddress          As Long

92     Select Case EndpointField
    Case LOCAL_HOST, LOCAL_HOST_IP, LOCAL_PORT

        'If the info of a local end-point of the connection is
        'requested, call the getsockname Winsock API function
93         lngReturnValue = getsockname(lngSocket, udtSocketAddress, LenB(udtSocketAddress))
94     Case REMOTE_HOST, REMOTE_HOST_IP, REMOTE_PORT

        'If the info of a remote end-point of the connection is
        'requested, call the getpeername Winsock API function
95         lngReturnValue = getpeername(lngSocket, udtSocketAddress, LenB(udtSocketAddress))
96     End Select

97     If lngReturnValue = 0 Then
        'If no errors occurred, the getsockname or getpeername function returns 0.

98         Select Case EndpointField
        Case LOCAL_PORT, REMOTE_PORT
            'Get the port number from the sin_port field and convert the byte ordering
99             GetIPEndPointField = IntegerToUnsigned(ntohs(udtSocketAddress.sin_port))

100         Case LOCAL_HOST_IP, REMOTE_HOST_IP

            'Get pointer to the string that contains the IP address
101             lngPtrToAddress = inet_ntoa(udtSocketAddress.sin_addr)

            'Retrieve that string by the pointer
102             GetIPEndPointField = StringFromPointer(lngPtrToAddress)
103         Case LOCAL_HOST, REMOTE_HOST

            'The same procedure as for an IP address only using GetHostNameByAddress
104             lngPtrToAddress = inet_ntoa(udtSocketAddress.sin_addr)
105             strIPAddress = StringFromPointer(lngPtrToAddress)
106             lngAddress = inet_addr(strIPAddress)
107             GetIPEndPointField = GetHostNameByAddress(lngAddress)

108         End Select
        'An error occured
109     Else
110         GetIPEndPointField = SOCKET_ERROR
111     End If

112 Exit Function

113 ErrorHandler:
114     Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in GetIPEndPointField on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Function

' Get the hostname of an endpoint (client or server).

Private Function GetHostNameByAddress(lngInetAdr As Long) As String

115     On Error GoTo ErrorHandler

116 Dim lngPtrHostEnt As Long
117 Dim udtHostEnt    As HOSTENT
118 Dim strHostName   As String

    'Get the pointer to the HOSTENT structure
119     lngPtrHostEnt = gethostbyaddr(lngInetAdr, 4, AF_INET)

    'Copy data into the HOSTENT structure
120     RtlMoveMemory udtHostEnt, ByVal lngPtrHostEnt, LenB(udtHostEnt)

    'Prepare the buffer to receive a string
121     strHostName = String$(256, 0)

    'Copy the host name into the strHostName variable
122     RtlMoveMemory ByVal strHostName, ByVal udtHostEnt.hName, 256

    'Cut received string by first chr(0) character
123     GetHostNameByAddress = Left$(strHostName, InStr(1, strHostName, Chr(0)) - 1)

124 Exit Function

125 ErrorHandler:
126     Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in GetHostNameByAddress on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Function

Function GetHostByNameAlias(ByVal hostname As String) As Long

127     On Error Resume Next
128     Dim phe As Long
129     Dim heDestHost As HOSTENT
130     Dim addrList As Long
131     Dim retIP As Long
132         retIP = inet_addr(hostname)
133         If retIP = INADDR_NONE Then
134             phe = gethostbyname(hostname)
135             If phe <> 0 Then
136                 CopyMemory heDestHost, ByVal phe, hostent_size
137                 CopyMemory addrList, ByVal heDestHost.hAddrList, 4
138                 CopyMemory retIP, ByVal addrList, heDestHost.hLength
139             Else
140                 retIP = INADDR_NONE
141             End If
142         End If
143         GetHostByNameAlias = retIP
144         If Err Then GetHostByNameAlias = INADDR_NONE

End Function

' Get the error description of a socket error.

Public Function GetErrorDescription(ByVal lngErrorCode As Long) As String

145     On Error GoTo ErrorHandler

146 Dim strDesc As String

147     Select Case lngErrorCode
    Case WSAEACCES
148         strDesc = "Permission denied."
149     Case WSAEADDRINUSE
150         strDesc = "Address already in use."
151     Case WSAEADDRNOTAVAIL
152         strDesc = "Cannot assign requested address."
153     Case WSAEAFNOSUPPORT
154         strDesc = "Address family not supported by protocol family."
155     Case WSAEALREADY
156         strDesc = "Operation already in progress."
157     Case WSAECONNABORTED
158         strDesc = "Software caused connection abort."
159     Case WSAECONNREFUSED
160         strDesc = "Connection refused."
161     Case WSAECONNRESET
162         strDesc = "Connection reset by peer."
163     Case WSAEDESTADDRREQ
164         strDesc = "Destination address required."
165     Case WSAEFAULT
166         strDesc = "Bad address."
167     Case WSAEHOSTDOWN
168         strDesc = "Host is down."
169     Case WSAEHOSTUNREACH
170         strDesc = "No route to host."
171     Case WSAEINPROGRESS
172         strDesc = "Operation now in progress."
173     Case WSAEINTR
174         strDesc = "Interrupted function call."
175     Case WSAEINVAL
176         strDesc = "Invalid argument."
177     Case WSAEISCONN
178         strDesc = "Socket is already connected."
179     Case WSAEMFILE
180         strDesc = "Too many open files."
181     Case WSAEMSGSIZE
182         strDesc = "Message too long."
183     Case WSAENETDOWN
184         strDesc = "Network is down."
185     Case WSAENETRESET
186         strDesc = "Network dropped connection on reset."
187     Case WSAENETUNREACH
188         strDesc = "Network is unreachable."
189     Case WSAENOBUFS
190         strDesc = "No buffer space available."
191     Case WSAENOPROTOOPT
192         strDesc = "Bad protocol option."
193     Case WSAENOTCONN
194         strDesc = "Socket is not connected."
195     Case WSAENOTSOCK
196         strDesc = "Socket operation on nonsocket."
197     Case WSAEOPNOTSUPP
198         strDesc = "Operation not supported."
199     Case WSAEPFNOSUPPORT
200         strDesc = "Protocol family not supported."
201     Case WSAEPROCLIM
202         strDesc = "Too many processes."
203     Case WSAEPROTONOSUPPORT
204         strDesc = "Protocol not supported."
205     Case WSAEPROTOTYPE
206         strDesc = "Protocol wrong type for socket."
207     Case WSAESHUTDOWN
208         strDesc = "Cannot send after socket shutdown."
209     Case WSAESOCKTNOSUPPORT
210         strDesc = "Socket type not supported."
211     Case WSAETIMEDOUT
212         strDesc = "Connection timed out."
213     Case WSATYPE_NOT_FOUND
214         strDesc = "Class type not found."
215     Case WSAEWOULDBLOCK
216         strDesc = "Resource temporarily unavailable."
217     Case WSAHOST_NOT_FOUND
218         strDesc = "Host not found."
219     Case WSANOTINITIALISED
220         strDesc = "Successful WSAStartup not yet performed."
221     Case WSANO_DATA
222         strDesc = "Valid name, no data record of requested type."
223     Case WSANO_RECOVERY
224         strDesc = "This is a nonrecoverable error."
225     Case WSASYSCALLFAILURE
226         strDesc = "System call failure."
227     Case WSASYSNOTREADY
228         strDesc = "Network subsystem is unavailable."
229     Case WSATRY_AGAIN
230         strDesc = "Nonauthoritative host not found."
231     Case WSAVERNOTSUPPORTED
232         strDesc = "Winsock.dll version out of range."
233     Case WSAEDISCON
234         strDesc = "Graceful shutdown in progress."
235     Case Else
236         strDesc = "Unknown error."
237     End Select

238     GetErrorDescription = strDesc

239 Exit Function

240 ErrorHandler:
241     Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in GetErrorDescription on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Function

Public Sub SocketsCleanup()

242     On Error GoTo ErrorHandler

243     WSACleanup

244 Exit Sub

245 ErrorHandler:
246     Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in SocketsCleanup on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Sub

Public Function SocketsInitialize() As Boolean

247     On Error GoTo ErrorHandler
248 Dim WSAD            As WSAData

249     SocketsInitialize = False

250     If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then Exit Function
251     If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then Exit Function

252     SocketsInitialize = True

253 Exit Function

254 ErrorHandler:
255     Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in SocketsInitialize on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Function

Public Function SetSockLinger(ByVal SockNum&, ByVal OnOff%, ByVal LingerTime%) As Long

256 Dim Linger As LingerType

257     Linger.l_onoff = OnOff
258     Linger.l_linger = LingerTime

259     If WSAsetsockopt(SockNum, SOL_SOCKET, SO_LINGER, Linger, 4) Then
260         Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in SetSockLinger on line " & Erl() & " triggered by " & Err.Source & vbCrLf & WSAGetLastError()
261     Else
262         If WSAgetsockopt(SockNum, SOL_SOCKET, SO_LINGER, Linger, 4) Then
263             Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in SetSockLinger on line " & Erl() & " triggered by " & Err.Source & vbCrLf & WSAGetLastError()
264         End If
265     End If

End Function

Function getascip(ByVal inn As Long) As String

266     On Error Resume Next
267     Dim lpStr&
268     Dim nStr&
269     Dim retString$
270         retString = String(32, 0)
271         lpStr = inet_ntoa(inn)
272         If lpStr = 0 Then
273             getascip = "255.255.255.255"
274             Exit Function
275         End If
276         nStr = lstrlen(lpStr)
277         If nStr > 32 Then nStr = 32
278         CopyMemory ByVal retString, ByVal lpStr, nStr
279         retString = Left(retString, nStr)
280         getascip = retString
281         If Err Then getascip = "255.255.255.255"

End Function

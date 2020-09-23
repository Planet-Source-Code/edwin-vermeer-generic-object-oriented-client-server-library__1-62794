VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenericClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' Generic Client Server class objects 1.1
'
' Copyright 2005, E.V.I.C.T. B.V.
' Website: http:\\www.evict.nl
' Support: mailto:evict@vermeer.nl
'
'Purpose:
' The GenericClient object can be used to create any sort of client application.
' This object will hide the difficulties in the Winsock API calls for you.
' You will probably write your own protocol handler (like SMTP, POP or ...) in the OnDataArive event.
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
' feb 2005 : Added extra error checking in the Class_Initialize and the Connect method

Option Explicit

'Some variables that are used to store the state of this object.
Private m_intConnectionPort As Long    'Just remember the port that we connected to
Private m_lngConnectionSocket As Long     'Just remember the socket where the connection is set up for
Private m_objConnection As Connection  'When we are connected to the server this object will be created for handeling this connection.
Private m_strReturnIpPort As String
'We need a memory block for resolving the hostname. these variables will be used for handeling those.
Private lngMemoryHandle    As Long      'Handle of the allocated memory block object
Private lngMemoryPointer   As Long      'Address of the memory block

'We are going to subclass a generated window and add 2 new message numbers for the communication
Private objSuperClass As Subclass               'This is the subclass handler
Implements WinSubHook.iSubclass                  'This is the interface to the subclass handler

Private lngWinsockMessage As Long          'Winsock resolve host message
Private lngResolveHostMessage As Long   'General Winsock message
Private lngWindowHandle As Long            'Window handle

'Purpose: This event will be called when the connection is closed by the server or the client itself.
Public Event OnClose()
'Purpose: This event will be called when the connect method executed successfully.
Public Event OnConnect()
'Purpose: This event will be called when data arived. After this you have to call the .Connection.Recv method to get the data.
Public Event OnDataArrive()
'Purpose: This event will be called whenever there is a problem.
Public Event OnError(lngRetCode As Long, strDescription As String)

'Purpose:
' The moment an instance of this object is created we can set up a new socket for communication.

Private Sub Class_Initialize()

1     On Error GoTo ErrorHandler
2 Dim WSAInfo As WSAData
3 Dim lngReturn As Long
    'Create a blank, invisible window
4     lngWindowHandle = CreateWindowEx(0&, "STATIC", "SOCKET_WINDOW", 0&, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&)

    'Create 2 unique message numbers for our messages
5     lngWinsockMessage = RegisterWindowMessage(App.EXEName & Trim(Str(lngWindowHandle)) & ".WinsockMessage")
6     lngResolveHostMessage = RegisterWindowMessage(App.EXEName & Trim(Str(lngWindowHandle)) & ".ResolveHostMessage")

    'Start up the subclassing for the new blank invisible window. This will only be done for the 2 new messages
7     Set objSuperClass = New Subclass        'Create a cSuperClass instance
8     With objSuperClass
9         .AddMsg lngWinsockMessage, MSG_AFTER
10         .AddMsg lngResolveHostMessage, MSG_AFTER
11         .Subclass lngWindowHandle, Me
12     End With
    'Later on when we start the connection we will tell the created winsock to
    'post it's messages to the above subclassed window.

    'Start version 1.1 of the Microsoft Winsock API
13     lngReturn = WSAStartup(WS_VERSION_REQD, WSAInfo)
14     If lngReturn <> 0 Then
15         Err.Raise vbObjectError Or lngReturn, "GenericClient_Initialize", "GenericClient :: Error in Class_Initialize on line 12 triggered by WSAStartup   (" & lngReturn & ")" & vbCrLf & "The 32-bit Windows Socket is not responding."
16     End If

17     If WSAInfo.iMaxSockets < MIN_SOCKETS_REQD Then
18         Err.Raise vbObjectError Or lngReturn, "GenericClient_Initialize", "GenericClient :: Error in Class_Initialize on line 12 triggered by WSAStartup" & vbCrLf & "This application requires a minimum of " & CStr(MIN_SOCKETS_REQD) & " supported sockets."
19     End If
20     If LoByte(WSAInfo.wVersion) < WS_VERSION_MAJOR Or (LoByte(WSAInfo.wVersion) = WS_VERSION_MAJOR And HiByte(WSAInfo.wVersion) < WS_VERSION_MINOR) Then
21         Err.Raise vbObjectError Or lngReturn, "GenericClient_Initialize", "GenericClient :: Error in Class_Initialize on line 12 triggered by WSAStartup" & vbCrLf & "Sockets version " & CStr(LoByte(WSAInfo.wVersion)) & "." & CStr(HiByte(WSAInfo.wVersion)) & " is not supported by 32-bit Windows Sockets."
22     End If

    'Since we only need one and the chance is big we are going to connect, just create the connection object now
23     Set m_objConnection = New Connection

24 Exit Sub

25 ErrorHandler:
26     Err.Raise vbObjectError Or Err, "GenericClient", "GenericClient :: Error in Class_Initialize on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Sub

'Purpose:
' When this object is closed we just make sure that we close the socket, clean up the connection and stop the subclassing.

Private Sub Class_Terminate()

27     On Error GoTo ErrorHandler

    'clean up the connection object
28     m_objConnection.CloseSocket
29     Set m_objConnection = Nothing

    'Cleanup Winsock

30     If WSAIsBlocking() Then WSACancelBlockingCall
31     Debug.Print "WSACleanup"
32     WSACleanup

    'Stop subclassing
33     objSuperClass.UnSubclass
34     Set objSuperClass = Nothing  'Destroy the subclasser
    'Also remove the window
35     DestroyWindow lngWindowHandle

36 Exit Sub

37 ErrorHandler:
38     Err.Raise vbObjectError Or Err, "GenericClient", "GenericClient :: Error in Class_Terminate on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Sub

'Purpose:
' This method will do the actual connect to the server.
' Before sending data make sure that te onConnect event was fired (and onDisconnect not).
' This method will return before the actuall connection was set up.

Public Function Connect(ByVal strRemoteHost As String, ByVal intRemotePort As Long) As Long

39     On Error GoTo ErrorHandler
40 Dim lngSocket As Long
41 Dim sockin As SOCKADDR_IN

    ' test the port usage
42     sockin.sin_family = AF_INET
43     sockin.sin_port = htons(intRemotePort)
44     If sockin.sin_port = INVALID_SOCKET Then
45         Err.Raise vbObjectError Or sockin.sin_port, "GenericClient_Connect", "GenericClient :: Error in Connect on line " & Erl() & " triggered by htons" & vbCrLf & "The Windows Sockets htons function could not converts a u_short from host to TCP/IP network byte order."
46     End If

    ' Test the remote hostname
47     sockin.sin_addr = GetHostByNameAlias(strRemoteHost)
48     If sockin.sin_addr = INADDR_NONE Then
49         Err.Raise vbObjectError Or INADDR_NONE, "GenericClient_Connect", "GenericClient :: Error in Connect on line " & Erl() & " triggered by htons" & vbCrLf & "The Windows Sockets gethostbyname function could not retrieve the host information corresponding to a host name from a host database."
50     End If
51     m_strReturnIpPort = getascip$(sockin.sin_addr) & ":" & ntohs(sockin.sin_port)

    'Create a new socket
52     lngSocket = Socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
53     m_objConnection.Socket = lngSocket
54     If lngSocket = INVALID_SOCKET Then
55         lngSocket = WSAGetLastError()
56         If lngSocket > 0 Then Err.Raise lngSocket + vbObjectError, "Socket", "Winsock error " & lngSocket & " wile connecting. " & vbCrLf & GetErrorDescription(lngSocket)
57     End If

    ' Test the sockLinger
58     If SetSockLinger(lngSocket, 1, 0) = SOCKET_ERROR Then
59         Err.Raise vbObjectError Or SOCKET_ERROR, "GenericClient_Connect", "GenericClient :: Error in Connect on line " & Erl() & " triggered by " & Err.Source & vbCrLf & "Error getting linger info."
60     End If

    'Let the socket send the messages to the subclassed window handle (see class_initialize)
61     WSAAsyncSelect lngSocket, lngWindowHandle, lngWinsockMessage, FD_CONNECT Or FD_READ Or FD_CLOSE Or FD_ACCEPT
62     Connect = lngSocket
63     If lngSocket = SOCKET_ERROR Then
64         lngSocket = WSAGetLastError()
65         If lngSocket > 0 Then Err.Raise lngSocket + vbObjectError, "WSAAsyncSelect", "Winsock error " & lngSocket & " wile connecting. " & vbCrLf & GetErrorDescription(lngSocket)
66     End If

67     If WSAConnect(lngSocket, sockin, sockaddr_size) <> -1 Then
68         Err.Raise vbObjectError Or 1, "GenericClient_Connect", "GenericClient :: Error in Connect on line " & Erl() & " triggered by " & Err.Source & vbCrLf & "Error connecting to remote host."
69     End If

    'Resolve the host name and if that was successful then give the socket a connect command.
70     ResolveHost strRemoteHost

    'Save the connection parameters
71     m_intConnectionPort = intRemotePort
72     m_lngConnectionSocket = lngSocket

73 Exit Function

74 ErrorHandler:
75     Err.Raise vbObjectError Or Err, "GenericClient", "GenericClient :: Error in Connect on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Function

' We can access the connection object from code with this property.
' This will let you use a sintax like GenericClient.Connection.Send

Public Property Get Connection() As Connection

76     On Error GoTo ErrorHandler

77     Set Connection = m_objConnection

78 Exit Property

79 ErrorHandler:
80     Err.Raise vbObjectError Or Err, "GenericClient", "GenericClient :: Error in Connection on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Property

' This is where the winsock messages are handeled.

Private Sub iSubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

81     On Error GoTo ErrorHandler
    'wParam is the socket
    'lParam is the winsock message
    'uMsg is the windows message type
82 Dim lngErrorCode As Long

    'was an error passed on
83     lngErrorCode = HiWord(lParam)
84     If lngErrorCode > 0 Then
85         RaiseEvent OnError(lngErrorCode, GetErrorDescription(lngErrorCode))
86     End If

    'Was it a winsock or resolvehost message?
87     Select Case uMsg
    Case lngWinsockMessage
        'What winsock message was posted
88         Select Case lParam
        Case FD_CONNECT     'The connection is established
89             RaiseEvent OnConnect
90             Connection.IsConnected = True
91             lReturn = 0
92         Case FD_WRITE
            ' We are ready to write.
93         Case FD_READ        'We are ready to read from the winsock buffer

            ' We needed this as a solution to handle faster incomming data than what we can process
94 Static intBufferCounter As Long
95             intBufferCounter = intBufferCounter + 1
96             If intBufferCounter > 1 Then Exit Sub
97 ContinueWithBuffer:
98             RaiseEvent OnDataArrive
99             intBufferCounter = intBufferCounter - 1
100             If intBufferCounter > 0 Then GoTo ContinueWithBuffer

101             lReturn = 0
102         Case FD_CLOSE       'The other end point wishes to close the connection
103             RaiseEvent OnClose
104             lReturn = 0
105         End Select
106     Case lngResolveHostMessage
        'A resolve message was posted. This only happens if we connect to a name instead of an IP adress.
107 Dim udtHost           As HOSTENT
108 Dim lngIpAddrPtr      As Long
109 Dim lngHostAddress    As Long
110         If Not lngErrorCode > 0 Then
            'Extract the host name from the memory block
111             RtlMoveMemory udtHost, ByVal lngMemoryPointer, Len(udtHost)
112             RtlMoveMemory lngIpAddrPtr, ByVal udtHost.hAddrList, 4
113             RtlMoveMemory lngHostAddress, ByVal lngIpAddrPtr, 4
            'Free the allocated memory block
114             Call GlobalUnlock(lngMemoryHandle)
115             Call GlobalFree(lngMemoryHandle)
116         Else
117             lngHostAddress = INADDR_NONE
118         End If
        'Now connect to the resolved host.
119         ResolveHostMessage lngHostAddress
120     End Select

121 Exit Sub

122 ErrorHandler:
123     Err.Raise vbObjectError Or Err, "GenericClient", "GenericClient :: Error in iSubclass_After on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Sub

' When you subclass you can receive messages before or after the original windows handler was called.

Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, hWnd As Long, uMsg As WinSubHook.eMsg, wParam As Long, lParam As Long)

' We only use the after.

End Sub

'Connect to the host that was provided. If it was a name instead of an IP then it will be resolved first.

Private Sub ResolveHost(strHostName As String)

124     On Error GoTo ErrorHandler
125 Dim lngAddress As Long

    'Try and resolve the address. This will work if it was an IP we were given
126     lngAddress = inet_addr(strHostName)

    'We were unable to resolve it so we will have to go for the long way
127     If lngAddress = INADDR_NONE Then
        'Allocate 1Kb of fixed memory
128         lngMemoryHandle = GlobalAlloc(GMEM_FIXED, 1024)

129         If lngMemoryHandle > 0 Then
            'Lock the memory block just to get the address
130             lngMemoryPointer = GlobalLock(lngMemoryHandle)

131             If lngMemoryPointer = 0 Then
                'Memory allocation error
132                 Call GlobalFree(lngMemoryHandle)
133                 Exit Sub
134             Else
                'Unlock the memory block
135                 GlobalUnlock (lngMemoryHandle)
136             End If
137         Else
            'Memory allocation error
138             Exit Sub
139         End If

        'Get the host by the name. This is an Asynchroneous call. This means
        'that the call will not freeze the app. It will post a message
        'to the WindowProc when it has finished.
140         WSAAsyncGetHostByName lngWindowHandle, lngResolveHostMessage, strHostName, ByVal lngMemoryPointer, 1024
141     Else
        'A valid IP adress was suplied as a hostname so we can connect immediately.
142         ResolveHostMessage (lngAddress)
143     End If

144 Exit Sub

145 ErrorHandler:
146     Err.Raise vbObjectError Or Err, "GenericClient", "GenericClient :: Error in ResolveHost on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Sub

' Give a connect message to the winsock.

Private Sub ResolveHostMessage(lngAddress As Long)

147     On Error GoTo ErrorHandler
148 Dim udtSocketAddress As SOCKADDR_IN

    ' We will only connect if we have a valid IP adress (converted to a long)
149     If lngAddress <> INADDR_NONE Then
150         With udtSocketAddress
151             .sin_addr = lngAddress
152             .sin_port = htons(UnsignedToInteger(CLng(m_intConnectionPort)))
153             .sin_family = AF_INET
154         End With
        'Call the connect API function
155         Call WSAConnect(m_lngConnectionSocket, udtSocketAddress, LenB(udtSocketAddress))
156     End If

157 Exit Sub

158 ErrorHandler:
159     Err.Raise vbObjectError Or Err, "GenericClient", "GenericClient :: Error in ResolveHostMesssage on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Sub

Public Function HiByte(ByVal wParam As Integer)

160     HiByte = wParam \ &H100 And &HFF&

End Function

Public Function LoByte(ByVal wParam As Integer)

161     LoByte = wParam And &HFF&

End Function

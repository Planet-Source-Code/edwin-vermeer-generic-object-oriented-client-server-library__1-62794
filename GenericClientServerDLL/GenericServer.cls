VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenericServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' Generic Clinet Server class objects 1.1
'
' Copyright 2005, E.V.I.C.T. B.V.
' Website: http:\\www.evict.nl
' Support: mailto:evict@vermeer.nl
'
'Purpose:
' The GenericServer object can be used to create any sort of server application.
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

Option Explicit

'Some variables that are used to store the state of this object.
Private m_lngSocket As Long               'Just remember the socket where we listen on
Private m_objConnection As New Collection 'When a client is connected to this server then a new connection object will be added to this collection for handeling this connection.
Attribute m_objConnection.VB_VarHelpID = -1
Private m_intConnectionPort As Long    'Just remember the port where we listen on

' We are going to subclass a generated window and add 2 new message numbers for the communication
Private objSuperClass As Subclass               'This is the subclass handler
Implements WinSubHook.iSubclass           'This is the interface to the subclass handler
Private WinsockMessage As Long            'Winsock resolve host message
Private WindowHandle As Long              'Window handle

'Purpose: This event will be called when the connection is closed by the server or the client itself.
Public Event OnClose(lngSocket As Long)
'Purpose: This event will be called when the connect method executed successfully.
Public Event OnConnectRequest(lngSocket As Long)
'Purpose: This event will be called when data arived. After this you have to call the .Connection.Recv method to get the data.
Public Event OnDataArrive(lngSocket As Long)
'Purpose: This event will be called whenever there is a problem.
Public Event OnError(lngRetCode As Long, strDescription As String)

'Purpose:
' The moment an instance of this object is created we can set up a new socket for communication.

Private Sub Class_Initialize()

1     On Error GoTo ErrorHandler
2 Dim WSAInfo As WSAData

    'Create a blank, invisible window
3     WindowHandle = CreateWindowEx(0&, "STATIC", "SOCKET_WINDOW", 0&, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&)

    'Create a unique message number for our messages
4     WinsockMessage = RegisterWindowMessage(App.EXEName & Trim(Str(WindowHandle)) & ".WinsockMessage")

    'Start up the subclassing
5     Set objSuperClass = New Subclass        'Create a cSuperClass instance
6     With objSuperClass
7         .AddMsg WinsockMessage, MSG_AFTER
8         .Subclass WindowHandle, Me
9     End With

    'Start version 1.1 of the Microsoft Winsock API
10     WSAStartup &H101, WSAInfo

11 Exit Sub

12 ErrorHandler:
13     Err.Raise vbObjectError Or Err, "GenericServer", "GenericServer :: Error in Class_Initialize on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Sub

' When this object is closed we just make sure that we close the socket, clean up all the connection and stop the subclassing.

Private Sub Class_Terminate()

14     On Error GoTo ErrorHandler

    'Close all clients that are connected
15     CloseAll

    'Cleanup Winsock
16     If WSAIsBlocking() Then WSACancelBlockingCall
17     WSACleanup

    'Stop subclassing
18     objSuperClass.UnSubclass
19     Set objSuperClass = Nothing  'Destroy the subclasser

    'Also remove the window
20     DestroyWindow WindowHandle

21 Exit Sub

22 ErrorHandler:
23     Err.Raise vbObjectError Or Err, "GenericServer", "GenericServer :: Error in Class_Terminate on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Sub

' This method will set the socket in listening mode.
' The moment a client requests a connection the event OnConnectRequest will be trigered.
' You must then call the .Accept method to accept this connection.

Public Function Listen(lngLocalPort As Long) As Long

24     On Error GoTo ErrorHandler
25 Dim udtSocketAddress As SOCKADDR_IN
26 Dim lngReturnValue   As Long
27 Dim lngAddress       As Long

    'Create a new socket
28     m_lngSocket = Socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)

    'Let the socket send the messages to the subclassed window handle
29     WSAAsyncSelect m_lngSocket, WindowHandle, WinsockMessage, FD_CONNECT Or FD_READ Or FD_CLOSE Or FD_ACCEPT
30     Listen = m_lngSocket
31     m_intConnectionPort = lngLocalPort

    'If Not m_lngSocket > 0 Then Exit Function 'Check the socket handle
    'If Not lngLocalPort > 0 Then Exit Function 'Check the port number
32     With udtSocketAddress
33         .sin_addr = inet_addr("0.0.0.0")                    'We will accept any IP
34         .sin_port = htons(UnsignedToInteger(lngLocalPort))  'The port to listen on
35         .sin_family = AF_INET                               'The internet family (TCP/IP, UDP etc)
36     End With

37     Call WSABind(m_lngSocket, udtSocketAddress, LenB(udtSocketAddress))

38     If WSAListen(m_lngSocket, SOMAXCONN) <> 0 Then
39         RaiseEvent OnError(vbObjectError + 1, "Could not start listening on port " & lngLocalPort)
40     End If

41 Exit Function

42 ErrorHandler:
43     Err.Raise vbObjectError Or Err, "GenericServer", "GenericServer :: Error in Listen on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Function

' You have to call this method to accept a client that wants to connect to this server.

Public Function Accept(ByVal lngSocketHandle As Long) As Long

44     On Error GoTo ErrorHandler
45 Dim udtSocketAddress As SOCKADDR_IN

    'Call the accept Winsock API function in order to create a new socket
46     Accept = WSAAccept(lngSocketHandle, udtSocketAddress, LenB(udtSocketAddress))

47 Dim m_C As New Connection
48     m_C.Socket = Accept
49     m_objConnection.Add m_C

50 Exit Function

51 ErrorHandler:
52     Err.Raise vbObjectError Or Err, "GenericServer", "GenericServer :: Error in Accept on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Function

'Make sure that all clients are disconnected.

Public Function CloseAll()

53     On Error GoTo ErrorHandler
54 Dim l As Long

    ' Close all connections
55     If m_objConnection.Count > 0 Then
56         For l = m_objConnection.Count To 1 Step -1
57             m_objConnection(l).CloseSocket
58             m_objConnection.Remove l
59         Next l
60     End If

    ' Close the listening socket
61     WSACloseSocket m_lngSocket
62     m_lngSocket = 0

63 Exit Function

64 ErrorHandler:
65     Err.Raise vbObjectError Or Err, "GenericServer", "GenericServer :: Error in CloseAll on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Function

' This property can be used to access the connection collection.<BR><BR>
' For instance if you want to send data to a specific client then you can call GenericServer.Connection(lngConnection).Send<BR>
' The variable lngConnection can be the socket number or the negative counter ID<BR>
' For instance if you have 1 connection with socket number 812 then Connection(812) = Connection(-1)<BR>

Public Property Get Connection(lngConnection As Long) As Connection

66     On Error GoTo ErrorHandler
67     If lngConnection < 0 Then
68         Set Connection = m_objConnection(-lngConnection)
69     Else
70 Dim o As Connection
71         For Each o In m_objConnection
72             If o.Socket = lngConnection Then
73                 Set Connection = o
74                 Exit For
75             End If
76         Next o
77     End If

78 Exit Property

79 ErrorHandler:
80     Err.Raise vbObjectError Or Err, "GenericServer", "GenericServer :: Error in Connection on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Property

' This will give you the socket number where the server is listening on.

Public Property Get ListenSocket() As Long

81     On Error GoTo ErrorHandler

82     ListenSocket = m_lngSocket

83 Exit Property

84 ErrorHandler:
85     Err.Raise vbObjectError Or Err, "GenericServer", "GenericServer :: Error in ListenSocket on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Property

' This property will return the number of connected clients.

Public Property Get ConnectionCount() As Long

86     On Error GoTo ErrorHandler

87     ConnectionCount = m_objConnection.Count

88 Exit Property

89 ErrorHandler:
90     Err.Raise vbObjectError Or Err, "GenericServer", "GenericServer :: Error in ConnectionCount on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Property

' This is where the winsock messages are handeled.

Private Sub iSubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As WinSubHook.eMsg, ByVal wParam As Long, ByVal lParam As Long)

91     On Error GoTo ErrorHandler
    'We only use the _After
    'wParam is the socket
    'lParam is the winsock message
    'uMsg is the windows message type
92 Dim lngErrorCode As Long

    'was an error passed on
93     lngErrorCode = HiWord(lParam)
94     If lngErrorCode > 0 Then
95         RaiseEvent OnError(lngErrorCode, GetErrorDescription(lngErrorCode))
96     End If

    'Was it a winsock message?
97     If uMsg = WinsockMessage Then
        'What winsock message
98         Select Case lParam
        Case FD_ACCEPT      'There is a connection request pending
99             RaiseEvent OnConnectRequest(wParam)
100             lReturn = 0
101         Case FD_READ        'We are ready to read from the winsock buffer

            ' We needed this as a solution to handle faster incomming data than what we can process
102 Static intBufferCounter As Long
103             intBufferCounter = intBufferCounter + 1
104             If intBufferCounter > 1 Then Exit Sub
105 ContinueWithBuffer:
106             Connection(wParam).IsConnected = True
107             RaiseEvent OnDataArrive(wParam)
108             intBufferCounter = intBufferCounter - 1
109             If intBufferCounter > 0 Then GoTo ContinueWithBuffer

110             lReturn = 0
111         Case FD_CLOSE       'The other end point wishes to close the connection
112             RaiseEvent OnClose(wParam)
113             lReturn = 0
114         Case Else
115             Debug.Print lParam
116         End Select
117     End If

118 Exit Sub

119 ErrorHandler:
120     Err.Raise vbObjectError Or Err, "GenericServer", "GenericServer :: Error in iSubclass_After on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Sub

' When you subclass you can receive messages before or after the original windows handler was called.

Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, hWnd As Long, uMsg As WinSubHook.eMsg, wParam As Long, lParam As Long)

'We only use the _After

End Sub

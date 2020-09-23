VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Connection"
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
' This object is created for a client when a connection is made or for a server for every connection that was made.
' You can get connection specific information or receive and send data.
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
' feb 2005 : Corrected the RecvByte and SendByte methods

Option Explicit

Dim m_myObject As Object 'used for a custom object reference.
Dim m_lngSocket As Long  'The socket number where we are connected to
Dim m_blnIsConnected As Boolean

'Purpose:
' You can get the socket number of this connection.

Public Property Get Socket() As Long

1     On Error GoTo ErrorHandler

2     Socket = m_lngSocket

3 Exit Property

4 ErrorHandler:
5     Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in Get Socket on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Property

' It can only be set by other code within this dll.

Friend Property Let Socket(lngSocket As Long)

6     On Error GoTo ErrorHandler

7     m_lngSocket = lngSocket

8 Exit Property

9 ErrorHandler:
10     Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in Let Socket on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Property

'Purpose:
' You can get or set a custom object. Remember that if you set this you also have to set it to nothing when you end your code. Otherwise this will leave an object reference open whicht will keep an instance of Site Skinner loaded.
' The main purpuse for this is so that you can attach an object of your own to every server connection that is established. This way you can easily access connection specific properties like username, connect duration and last activity.
' You can even add a Site Skinner client object to it so that you can create a proxy, gateway or port mapper with very litle code. Look at the Client Server sample to see how it can be used.

Public Property Get CustomObject() As Object

11     On Error GoTo ErrorHandler

12     Set CustomObject = m_myObject

13 Exit Property

14 ErrorHandler:
15     Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in Get CustomObject on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Property

Public Property Let CustomObject(yourObject As Object)

16     On Error GoTo ErrorHandler

17     Set m_myObject = yourObject

18 Exit Property

19 ErrorHandler:
20     Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in Let CustomObject on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Property

Public Sub SetCustomObject(yourObject As Object)

21     On Error GoTo ErrorHandler

22     Set m_myObject = yourObject

23 Exit Sub

24 ErrorHandler:
25     Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in Let SetCustomObject on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Sub

'Purpse:
' You can get the name of the client computer.

Public Function GetLocalHost() As String

26     On Error GoTo ErrorHandler
27 Dim lngSocket As Long

28     lngSocket = m_lngSocket
29     GetLocalHost = GetIPEndPointField(lngSocket, LOCAL_HOST)

30 Exit Function

31 ErrorHandler:
32     Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in GetLocalHost on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Function

'Purpse:
' You can get the IP adress of the client computer.

Public Function GetLocalIP() As String

33     On Error GoTo ErrorHandler
34 Dim lngSocket As Long

35     lngSocket = m_lngSocket
36     GetLocalIP = GetIPEndPointField(lngSocket, LOCAL_HOST_IP)

37 Exit Function

38 ErrorHandler:
39     Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in GetLocalIP on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Function

'Purpse:
' You can get the connection port on the client computer.

Public Function GetLocalPort() As Long

40     On Error GoTo ErrorHandler
41 Dim lngSocket As Long

42     lngSocket = m_lngSocket
43     GetLocalPort = GetIPEndPointField(lngSocket, LOCAL_PORT)

44 Exit Function

45 ErrorHandler:
46     Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in GetLocalPort on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Function

'Purpse:
' You can get the name of the remote computer.

Public Function GetRemoteHost() As String

47     On Error GoTo ErrorHandler
48 Dim lngSocket As Long

49     lngSocket = m_lngSocket
50     GetRemoteHost = GetIPEndPointField(lngSocket, REMOTE_HOST)

51 Exit Function

52 ErrorHandler:
53     Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in GetRemoteHost on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Function

'Purpse:
' You can get the IP adress of the remote computer.

Public Function GetRemoteIP() As String

54     On Error GoTo ErrorHandler
55 Dim lngSocket As Long

56     lngSocket = m_lngSocket
57     GetRemoteIP = GetIPEndPointField(lngSocket, REMOTE_HOST_IP)

58 Exit Function

59 ErrorHandler:
60     Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in GetRemoteIP on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Function

'Purpse:
' You can get the port number of the remote computer.

Public Function GetRemotePort() As Long

61     On Error GoTo ErrorHandler
62 Dim lngSocket As Long

63     lngSocket = m_lngSocket
64     GetRemotePort = GetIPEndPointField(lngSocket, REMOTE_PORT)

65 Exit Function

66 ErrorHandler:
67     Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in GetRemotePort on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Function

'Purpose:
' Send data to the remote computer.
' This will use a Unicode to byte array conversion

Public Function Send(strData As String) As Long

68     On Error GoTo ErrorHandler
69 Dim lngSocket As Long

70     lngSocket = m_lngSocket
71     Send = mSend(lngSocket, strData)

72 Exit Function

73 ErrorHandler:
74     Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in Send on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Function

'Purpose:
' After receiving an OnDataArive event we can get the data with this.
' This will use a byte array to Unicode conversion.

Public Function Recv(ByRef strBuffer As String) As Long

75     On Error GoTo ErrorHandler
76 Dim lngSocket As Long

77     lngSocket = m_lngSocket
78     Recv = mRecv(lngSocket, strBuffer)

79 Exit Function

80 ErrorHandler:
81     Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in Recv on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Function

'Purpose:
' Send data to the remote computer.
' This will be done without any unicode conversion.

Public Function SendByte(ByRef byteData() As Byte) As Long

82     On Error GoTo ErrorHandler
83 Dim lngSocket As Long

84     lngSocket = m_lngSocket
85     SendByte = mSendByte(lngSocket, byteData())

86 Exit Function

87 ErrorHandler:
88     Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in SendByte on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Function

'Purpose:
' After receiving an OnDataArive event we can get the data with this.
' This will be done without any unicode conversion.

Public Function RecvByte(ByRef byteData() As Byte) As Long

89     On Error GoTo ErrorHandler
90 Dim lngSocket As Long

91     lngSocket = m_lngSocket
92     RecvByte = mRecvByte(lngSocket, byteData())

93 Exit Function

94 ErrorHandler:
95     Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in Recv on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Function

'Purpose:
' Close the connection.

Public Sub CloseSocket()

96     m_blnIsConnected = False

97     On Error GoTo ErrorHandler
98     If m_lngSocket = 0 Then Exit Sub
99     Debug.Print "WSACloseSocket " & m_lngSocket
100     WSACloseSocket m_lngSocket
101     m_lngSocket = 0

102 Exit Sub

103 ErrorHandler:
104     Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in CloseSocket on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Sub

Private Sub Class_Initialize()

105     m_blnIsConnected = False

End Sub

Private Sub Class_Terminate()

106     If m_lngSocket <> 0 Then CloseSocket

End Sub

Public Property Get IsConnected() As Boolean

107     IsConnected = m_blnIsConnected

End Property

Friend Property Let IsConnected(blnConnected As Boolean)

108     m_blnIsConnected = blnConnected

End Property

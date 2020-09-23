VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PortMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Port Mapper class objects 1.0
'
' Copyright 2005, E.V.I.C.T. B.V.
' Website: http:\\www.evict.nl
' Support: mailto:evict@vermeer.nl
'
'Purpose:
' This class object will encaptule the complete functionality for a port mapper and
' proxy server. Everything is handeled by this class. You only have to create an
' instance of this class and call the ListenAsPortmapper or ListenAsProxy method.
' If you want to show what is happening then the only thing what you have to do
' is to listen to the events that are raised by this class.
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
' feb 2005 : Created as part of the portmapper demo

Option Explicit

Public WithEvents cServer  As GenericServer
Attribute cServer.VB_VarHelpID = -1
' The socket were we listen on
Dim lngSocket As Long
' For saving the server and port where we are going to connect to
Dim p_strRemoteHost As String
Dim p_strRemotePort As String
Dim p_blnAsProxy As Boolean
Dim DataBuffer() As Byte
Dim DataBufferConnect(7) As Byte

' Events from the server
Public Event OnCloseServer(lngSocket As Long)
Public Event OnConnectRequestServer(lngSocket As Long)
Public Event OnErrorServer(lngSocket As Long, lngRetCode As Long, strDescription As String)
Public Event OnDataAriveServer(lngSocket As Long, DataBuffer() As Byte)
' Events from the client
Public Event OnCloseClient(lngSocket As Long, lngSocketParent As Long)
Public Event OnConnectClient(lngSocket As Long, lngSocketParent As Long)
Public Event OnErrorClient(lngSocket As Long, lngRetCode As Long, strDescription As String, lngSocketParent As Long)
Public Event OnDataAriveClient(lngSocket As Long, DataBuffer() As Byte, lngSocketParent As Long)


Private Sub Class_Initialize()
    Set cServer = New GenericServer
End Sub

Private Sub Class_Terminate()
    cServer.CloseAll
    Set cServer = Nothing
End Sub


'------------------------------------------------------------------
' The 2 methods to start up a Proxy or a portmapper
'------------------------------------------------------------------

Public Sub ListenAsPortmapper(strLocalPort As String, strRemoteHost As String, strRemotePort As String)
     p_strRemoteHost = strRemoteHost
     p_strRemotePort = strRemotePort
     p_blnAsProxy = False
     lngSocket = cServer.Listen(CLng(strLocalPort))
End Sub

Public Sub ListenAsProxy(strLocalPort As String)
     p_blnAsProxy = True
     lngSocket = cServer.Listen(CLng(strLocalPort))
End Sub


'------------------------------------------------------------------
' Handeling the server events
'------------------------------------------------------------------

Private Sub cServer_OnClose(lngSocketX As Long)
     'The connection at the listning side is closed so we can close the connection
     cServer.Connection(lngSocketX).CustomObject.cClient.Connection.CloseSocket
     RaiseEvent OnCloseServer(lngSocketX)
End Sub

Private Sub cServer_OnConnectRequest(lngSocketX As Long)
Dim lngNewSocket As Long
    'Accept the connection and store the new socket handle
    lngNewSocket = cServer.Accept(lngSocketX)
    'Add a new instance of the GenericClient class to the new connection
    Dim cPortMapperConnection As New PortMapperConnection
    cPortMapperConnection.setParent Me, lngNewSocket
    cServer.Connection(lngNewSocket).SetCustomObject cPortMapperConnection
    'If we are a portmapper then we can connect otherwise we have to wait for the first data package
    If p_blnAsProxy = False Then
        cServer.Connection(lngNewSocket).CustomObject.cClient.Connect p_strRemoteHost, CLng(p_strRemotePort)
    End If
    RaiseEvent OnConnectRequestServer(lngNewSocket)
End Sub

Private Sub cServer_OnError(lngRetCode As Long, strDescription As String)
    RaiseEvent OnErrorServer(lngSocket, lngRetCode, strDescription)
End Sub

Private Sub cServer_OnDataArrive(lngSocketX As Long)
Dim strData As String
Dim lngLoop As Long
Dim lngDataLength As Long
ReDim DataBuffer(8168) As Byte
     'Recieve dhe data
     lngDataLength = cServer.Connection(lngSocketX).RecvByte(DataBuffer())
     If lngDataLength > 0 Then
         ReDim Preserve DataBuffer(lngDataLength - 1) As Byte
         If p_blnAsProxy = True Then
            If cServer.Connection(lngSocketX).CustomObject.cClient.Connection.IsConnected = False Then
               'For a proxy the first package contains the connection specs
              If DataBuffer(0) = 4 And DataBuffer(1) = 1 Then
                ' This is a Socks4 proxy connection
                p_strRemotePort = Str(CLng(DataBuffer(2)) * 256 + DataBuffer(3))
                p_strRemoteHost = DataBuffer(4) & "." & DataBuffer(5) & "." & DataBuffer(6) & "." & DataBuffer(7)
                Debug.Print "new socket = " & cServer.Connection(lngSocketX).CustomObject.cClient.Connect(p_strRemoteHost, CLng(p_strRemotePort))
                For lngLoop = 2 To 7
                    DataBufferConnect(lngLoop) = DataBuffer(lngLoop)
                Next
              Else
                ' This is not a Socks4 proxy connection. Assume a Http proxy connection
' not ready yet.
                p_strRemotePort = 80
                p_strRemoteHost = "www.microsoft.com"
                cServer.Connection(lngSocketX).CustomObject.cClient.Connect p_strRemoteHost, CLng(p_strRemotePort)
                lngDataLength = cServer.Connection(lngSocketX).CustomObject.cClient.Connection.SendByte(DataBuffer())
              End If
             RaiseEvent OnDataAriveServer(lngSocketX, DataBuffer())
            Else
               'For a proxy the 2nd and rest of the packages will just be send
               'The event is raised first so that we could manipulate the data before it is send on
               RaiseEvent OnDataAriveServer(lngSocketX, DataBuffer())
               lngDataLength = cServer.Connection(lngSocketX).CustomObject.cClient.Connection.SendByte(DataBuffer())
            End If
         Else
             'A portmapper will always just send the data
             'The event is raised first so that we could manipulate the data before it is send on
             RaiseEvent OnDataAriveServer(lngSocketX, DataBuffer())
             lngDataLength = cServer.Connection(lngSocketX).CustomObject.cClient.Connection.SendByte(DataBuffer())
         End If
     End If
End Sub



'------------------------------------------------------------------
' Client event mappings from the PortMapperConnection objects.
'------------------------------------------------------------------

Public Sub PortMapperConnection_onCloseClient(lngSocketX As Long, lngSocketParent As Long)
    RaiseEvent OnCloseClient(lngSocketX, lngSocketParent)
End Sub

Public Sub PortMapperConnection_onErrorClient(lngSocketX As Long, lngRetCode As Long, strDescription As String, lngSocketParent As Long)
    DataBufferConnect(0) = 4
    DataBufferConnect(1) = 91
    cServer.Connection(lngSocketParent).SendByte DataBufferConnect()
    RaiseEvent OnErrorClient(lngSocketX, lngRetCode, strDescription, lngSocketParent)
End Sub

Public Sub PortMapperConnection_onConnectClient(lngSocketX As Long, lngSocketParent As Long)
    DataBufferConnect(0) = 4
    DataBufferConnect(1) = 90
    cServer.Connection(lngSocketParent).SendByte DataBufferConnect()
    RaiseEvent OnConnectClient(lngSocketX, lngSocketParent)
End Sub

Public Sub PortMapperConnection_OnDataAriveClient(lngSocketX As Long, lngSocketParent As Long)
Dim strData As String
Dim lngDataLength As Long
Dim DataBuffer() As Byte
ReDim DataBuffer(8168) As Byte
    lngDataLength = cServer.Connection(lngSocketParent).CustomObject.cClient.Connection.RecvByte(DataBuffer())
    If lngDataLength > 0 Then
        ReDim Preserve DataBuffer(lngDataLength - 1) As Byte
        'The event is raised first so that we could manipulate the data before it is send on
        RaiseEvent OnDataAriveClient(lngSocketX, DataBuffer(), lngSocketParent)
        cServer.Connection(lngSocketParent).SendByte DataBuffer()
    End If
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PortMapperConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Port Mapper Connection class objects 1.0
'
' Copyright 2005, E.V.I.C.T. B.V.
' Website: http:\\www.evict.nl
' Support: mailto:evict@vermeer.nl
'
'Purpose:
' This is a wrapper class for the GenericClient class. This is neccessery so that
' we can get events for each client that we use in the portmapper class
' The lngSocket value will be passed on so that we can identify them.
' Events are simulated by calling a sub in the parent object.
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

Public WithEvents cClient  As GenericClient
Attribute cClient.VB_VarHelpID = -1
Dim cParent As PortMapper
Dim lngSocket As Long
Dim lngSocketParent As Long


' Initialize the GenericClient
Private Sub Class_Initialize()
    Set cClient = New GenericClient
End Sub

' Destroy the GenericClient
Private Sub Class_Terminate()
    cClient.Connection.CloseSocket
    Set cClient = Nothing
End Sub


' Get a reference to the server class and server connection.
Public Sub setParent(clsParent As PortMapper, lngParentsocket)
    lngSocketParent = lngParentsocket
    Set cParent = clsParent
End Sub


' Pass on all the client events.
Private Sub cClient_OnConnect()
    lngSocket = cClient.Connection.Socket
    cParent.PortMapperConnection_onConnectClient lngSocket, lngSocketParent
End Sub

Private Sub cClient_OnClose()
    cParent.PortMapperConnection_onCloseClient lngSocket, lngSocketParent
End Sub

Private Sub cClient_OnError(lngRetCode As Long, strDescription As String)
    cParent.PortMapperConnection_onErrorClient lngSocket, lngRetCode, strDescription, lngSocketParent
End Sub

Private Sub cClient_OnDataArrive()
    cParent.PortMapperConnection_OnDataAriveClient lngSocket, lngSocketParent
End Sub

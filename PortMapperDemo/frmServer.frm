VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Server"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   6120
      Width           =   8295
   End
   Begin VB.CommandButton cmdDisconnectAll 
      Caption         =   "Disconnect all and quit"
      Height          =   435
      Left            =   8760
      TabIndex        =   0
      Top             =   7560
      Width           =   2655
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1515
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   2672
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483641
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   4515
      Left            =   0
      TabIndex        =   2
      Top             =   1560
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   7964
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483641
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Purpose:
' This demo will show how to create multiple instances of a generic client and server.
' Created by Edwin Vermeer
' Website http://siteskinner.com

Public WithEvents cPortMapper As PortMapper
Attribute cPortMapper.VB_VarHelpID = -1




Private Sub Form_Load()

1     On Error GoTo ErrorHandler

2     Set cPortMapper = New PortMapper
    ' Put in the listview headers
3     ListView1.ColumnHeaders.Add 1, , "Socket Handle"
4     ListView1.ColumnHeaders.Add 2, , "Remote Host"
5     ListView1.ColumnHeaders.Add 3, , "Remote IP"
6     ListView1.ColumnHeaders.Add 4, , "Remote Port"
7     ListView1.ColumnHeaders.Add 5, , "Start time"
8     ListView1.ColumnHeaders.Add 6, , "Data in"
9     ListView1.ColumnHeaders.Add 7, , "Data out"
10     ListView1.ColumnHeaders.Add 8, , "Last communication"

     ListView2.ColumnHeaders.Add 1, , "in/out"
     ListView2.ColumnHeaders.Add 2, , "Time"
     ListView2.ColumnHeaders.Add 3, , "Size"
     ListView2.ColumnHeaders.Add 4, , "From Socket"
     ListView2.ColumnHeaders.Add 5, , "From Host"
     ListView2.ColumnHeaders.Add 6, , "From IP"
     ListView2.ColumnHeaders.Add 7, , "From Port"
     ListView2.ColumnHeaders.Add 8, , "To Socket"
     ListView2.ColumnHeaders.Add 9, , "To Host"
     ListView2.ColumnHeaders.Add 10, , "To IP"
     ListView2.ColumnHeaders.Add 11, , "To Port"
     ListView2.ColumnHeaders.Add 12, , "Data"

11 Exit Sub

12 ErrorHandler:
13     HandleTheException "frmServer :: Error in Form_Load() on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in Form_Load()", enumExceptionHandling_Exception

End Sub

'Purpose:
' Make sure that all clients are disconnected and unload the server object.

Private Sub Form_Unload(Cancel As Integer)

14     On Error GoTo ErrorHandler

    'Unload the portmapper class - This MUST be done
15     cPortMapper.cServer.CloseAll
16     Set cPortMapper = Nothing
17 Exit Sub

18 ErrorHandler:
19     HandleTheException "frmServer :: Error in Form_Unload() on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in Form_Unload()", enumExceptionHandling_Exception

End Sub

'Purpose:
' Just stop.
Private Sub cmdDisconnectAll_Click()

20     On Error GoTo ErrorHandler

21     Unload Me

22 Exit Sub

23 ErrorHandler:
24     HandleTheException "frmServer :: Error in cmdDisconnectAll_Click() on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in cmdDisconnectAll_Click()", enumExceptionHandling_Exception

End Sub

'Purpose:
' Set the server in listening mode on the specified port.

Public Sub Listen(strLocalPort As String, strRemoteHost As String, strRemotePort As String, blnAsProxy As Boolean)

    If blnAsProxy Then
       cPortMapper.ListenAsProxy strLocalPort
29     Me.Caption = "Proxy listening at port " & CLng(strLocalPort)
    Else
      cPortMapper.ListenAsPortmapper strLocalPort, strRemoteHost, strRemotePort
      Me.Caption = "Portmapper listening at port " & CLng(strLocalPort) & " and mapping it to " & strRemoteHost & " on port " & CLng(strRemotePort)
    End If

30 Exit Sub

31 ErrorHandler:
32     HandleTheException "frmServer :: Error in Listen(" & strLocalPort & ", " & strRemoteHost & ", " & strRemotePort & ") on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in Listen(" & strRemotePort & ")", enumExceptionHandling_Exception

End Sub


'Purpose:
' Whatever was set up as the active connection can not be active anymore.

Private Sub ClearData(lngSocketX As Long)

38     On Error GoTo ErrorHandler
39 Dim lngLoop As Long

    'Remove it from the list
40     For lngLoop = 1 To ListView1.ListItems.Count
41         If CLng(ListView1.ListItems(lngLoop)) = lngSocketX Then
42             ListView1.ListItems.Remove lngLoop
43             Exit For
44         End If
45     Next lngLoop

46 Exit Sub

47 ErrorHandler:
48     HandleTheException "frmServer :: Error in ClearData(" & lngSocketX & ") on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in ClearData(" & lngSocketX & ")", enumExceptionHandling_Exception

End Sub



'----------------------------------------------------------
' The Server events
'----------------------------------------------------------

Private Sub cPortMapper_OnConnectRequestServer(lngSocketX As Long)
Debug.Print "cPortMapper_OnConnectRequestServer " & lngSocketX
Dim ListHeader    As ListItem

     Set ListHeader = ListView1.ListItems.Add(, , lngSocketX)
     ListHeader.SubItems(1) = cPortMapper.cServer.Connection(lngSocketX).GetRemoteHost
     ListHeader.SubItems(2) = cPortMapper.cServer.Connection(lngSocketX).GetRemoteIP
     ListHeader.SubItems(3) = cPortMapper.cServer.Connection(lngSocketX).GetRemotePort
     ListHeader.SubItems(4) = Now
     ListHeader.SubItems(5) = 0
     ListHeader.SubItems(6) = 0
     ListHeader.SubItems(7) = Now
     Me.Caption = "Server " & cPortMapper.cServer.Connection(lngSocketX).GetLocalHost & " (" & cPortMapper.cServer.Connection(lngSocketX).GetLocalIP & ") is listening at port " & cPortMapper.cServer.Connection(lngSocketX).GetLocalPort

End Sub

Private Sub cPortMapper_OnCloseServer(lngSocketX As Long)
Debug.Print "cPortMapper_OnCloseServer " & lngSocketX
    ClearData lngSocketX
End Sub

Private Sub cPortMapper_OnErrorServer(lngSocketX As Long, lngRetCode As Long, strDescription As String)
Debug.Print "cPortMapper_OnErrorServer " & lngSocketX & ", " & lngRetCode & ", " & strDescription & vbCrLf
End Sub

Private Sub cPortMapper_OnDataAriveServer(lngSocketX As Long, DataBuffer() As Byte)
Debug.Print "cPortMapper_OnDataAriveServer " & lngSocketX & ", [" & UBound(DataBuffer()) & " bytes]"
Dim lngLoop As Long
     For lngLoop = 1 To ListView1.ListItems.Count
         If CLng(ListView1.ListItems(lngLoop)) = lngSocketX Then
             ListView1.ListItems(lngLoop).SubItems(5) = ListView1.ListItems(lngLoop).SubItems(5) + UBound(DataBuffer)
             ListView1.ListItems(lngLoop).SubItems(7) = Now
             Exit For
         End If
     Next lngLoop

     On Error Resume Next
     Dim ListHeader    As ListItem
     Set ListHeader = ListView2.ListItems.Add(, , lngSocketX)
     ListHeader.SubItems(0) = "OUT"
     ListHeader.SubItems(1) = Now
     ListHeader.SubItems(2) = UBound(DataBuffer())
     ListHeader.SubItems(3) = cPortMapper.cServer.Connection(lngSocketX).Socket
     ListHeader.SubItems(4) = cPortMapper.cServer.Connection(lngSocketX).GetRemoteHost
     ListHeader.SubItems(5) = cPortMapper.cServer.Connection(lngSocketX).GetRemoteIP
     ListHeader.SubItems(6) = cPortMapper.cServer.Connection(lngSocketX).GetRemotePort
     ListHeader.SubItems(7) = cPortMapper.cServer.Connection(lngSocketX).CustomObject.cClient.Connection.Socket
     ListHeader.SubItems(8) = cPortMapper.cServer.Connection(lngSocketX).CustomObject.cClient.Connection.GetRemoteHost
     ListHeader.SubItems(9) = cPortMapper.cServer.Connection(lngSocketX).CustomObject.cClient.Connection.GetRemoteIP
     ListHeader.SubItems(10) = cPortMapper.cServer.Connection(lngSocketX).CustomObject.cClient.Connection.GetRemotePort
     ListHeader.SubItems(11) = StrConv(DataBuffer(), vbUnicode)

End Sub


'----------------------------------------------------------
' The Client events
'----------------------------------------------------------

Private Sub cPortMapper_OnCloseClient(lngSocketX As Long, lngSocketParent As Long)
Debug.Print "cPortMapper_OnCloseClient " & lngSocketX & ", " & lngSocketParent
'
End Sub

Private Sub cPortMapper_OnConnectClient(lngSocketX As Long, lngSocketParent As Long)
Debug.Print "cPortMapper_OnConnectClient " & lngSocketX & ", " & lngSocketParent
'
End Sub

Private Sub cPortMapper_OnDataAriveClient(lngSocketX As Long, DataBuffer() As Byte, lngSocketParent As Long)
Debug.Print "cPortMapper_OnDataAriveClient " & lngSocketX & ", [" & UBound(DataBuffer()) & " bytes], " & lngSocketParent
' Left(StrConv(DataBuffer(), vbUnicode), 200)
     
     On Error Resume Next
     Dim ListHeader    As ListItem
     Set ListHeader = ListView2.ListItems.Add(, , lngSocketParent)
     ListHeader.SubItems(0) = "IN"
     ListHeader.SubItems(1) = Now
     ListHeader.SubItems(2) = UBound(DataBuffer())
     ListHeader.SubItems(3) = cPortMapper.cServer.Connection(lngSocketParent).CustomObject.cClient.Connection.Socket
     ListHeader.SubItems(4) = cPortMapper.cServer.Connection(lngSocketParent).CustomObject.cClient.Connection.GetRemoteHost
     ListHeader.SubItems(5) = cPortMapper.cServer.Connection(lngSocketParent).CustomObject.cClient.Connection.GetRemoteIP
     ListHeader.SubItems(6) = cPortMapper.cServer.Connection(lngSocketParent).CustomObject.cClient.Connection.GetRemotePort
     ListHeader.SubItems(7) = cPortMapper.cServer.Connection(lngSocketParent).Socket
     ListHeader.SubItems(8) = cPortMapper.cServer.Connection(lngSocketParent).GetRemoteHost
     ListHeader.SubItems(9) = cPortMapper.cServer.Connection(lngSocketParent).GetRemoteIP
     ListHeader.SubItems(10) = cPortMapper.cServer.Connection(lngSocketParent).GetRemotePort
     ListHeader.SubItems(11) = StrConv(DataBuffer(), vbUnicode)

End Sub

Private Sub cPortMapper_OnErrorClient(lngSocketX As Long, lngRetCode As Long, strDescription As String, lngSocketParent As Long)
Debug.Print "cPortMapper_OnErrorClient " & lngSocketX & ", " & lngRetCode & ", " & strDescription & ", " & lngSocketParent
'
End Sub



Private Sub ListView2_Click()
Text1 = ListView2.SelectedItem.SubItems(11)
End Sub

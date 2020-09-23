VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Server"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   11550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Selected Client"
      Height          =   3135
      Left            =   0
      TabIndex        =   1
      Top             =   2760
      Width           =   11535
      Begin VB.CommandButton cmdSendAll 
         Caption         =   "Send all"
         Height          =   315
         Left            =   7680
         TabIndex        =   13
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Height          =   315
         Left            =   7680
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtDataRecv 
         Height          =   1575
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   1440
         Width           =   8535
      End
      Begin VB.TextBox txtSendData 
         Height          =   1080
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   7455
      End
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "Disconnect"
         Height          =   315
         Left            =   7680
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblRemotePort 
         Caption         =   "Remote Port:"
         Height          =   330
         Left            =   8760
         TabIndex        =   11
         Top             =   2460
         Width           =   2640
      End
      Begin VB.Label lblRemoteIP 
         Caption         =   "Remote IP:"
         Height          =   330
         Left            =   8760
         TabIndex        =   10
         Top             =   2040
         Width           =   2640
      End
      Begin VB.Label lblRemoteHost 
         Caption         =   "Remote Host"
         Height          =   330
         Left            =   8760
         TabIndex        =   9
         Top             =   1620
         Width           =   2640
      End
      Begin VB.Label lblLocalPort 
         Caption         =   "Local Port:"
         Height          =   330
         Left            =   8760
         TabIndex        =   8
         Top             =   1200
         Width           =   2640
      End
      Begin VB.Label lblLocalIP 
         Caption         =   "Local IP:"
         Height          =   330
         Left            =   8760
         TabIndex        =   7
         Top             =   780
         Width           =   2640
      End
      Begin VB.Label lblLocalHost 
         Caption         =   "Local Host:"
         Height          =   330
         Left            =   8760
         TabIndex        =   6
         Top             =   360
         Width           =   2640
      End
   End
   Begin VB.CommandButton cmdDisconnectAll 
      Caption         =   "Disconnect all and quit"
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   2280
      Width           =   2655
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2235
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   3942
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
'
' Please read the license information and description in one of the GenericClientServer cls files
'
'Purpose:
' This demo will show how to create multiple instances of a generic client and server.

' We will handle a server.
Option Explicit
Private WithEvents cServer  As GenericServer
Attribute cServer.VB_VarHelpID = -1
' The socket were we listen on
Dim lngSocket As Long

'Purpose:
' Create a new instance of the server object.

Private Sub Form_Load()

1     On Error GoTo ErrorHandler

2     Set cServer = New GenericServer
    ' Put in the listview headers
3     ListView1.ColumnHeaders.Add 1, , "Socket Handle"
4     ListView1.ColumnHeaders.Add 2, , "Remote Host"
5     ListView1.ColumnHeaders.Add 3, , "Remote IP"
6     ListView1.ColumnHeaders.Add 4, , "Remote Port"
7     ListView1.ColumnHeaders.Add 5, , "Start time"
8     ListView1.ColumnHeaders.Add 6, , "Data in"
9     ListView1.ColumnHeaders.Add 7, , "Data out"
10     ListView1.ColumnHeaders.Add 8, , "Last communication"

11 Exit Sub

12 ErrorHandler:
13     HandleTheException "frmServer :: Error in Form_Load() on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in Form_Load()", enumExceptionHandling_Exception

End Sub

'Purpose:
' Make sure that all clients are disconnected and unload the server object.

Private Sub Form_Unload(Cancel As Integer)

14     On Error GoTo ErrorHandler

    'Unload the client class - This MUST be done
15     cServer.CloseAll
16     Set cServer = Nothing

17 Exit Sub

18 ErrorHandler:
19     HandleTheException "frmServer :: Error in Form_Unload() on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in Form_Unload()", enumExceptionHandling_Exception

End Sub

'Purpose:
' Disconnect the active (click on one in the list) client.

Private Sub cmdDisconnect_Click()

20     On Error GoTo ErrorHandler

    ' You have to specify which connection to close
21     If lngSocket = 0 Then
22         MsgBox "First you have to select a connection!", vbCritical, "Close connection"
23     Else
        'Close the socket
24         cServer.Connection(lngSocket).CloseSocket
        'Clear data of active connection
25         ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
26         ClearData lngSocket
27     End If

28 Exit Sub

29 ErrorHandler:
30     HandleTheException "frmServer :: Error in cmdDisconnect_Click() on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in cmdDisconnect_Click()", enumExceptionHandling_Exception

End Sub

'Purpose:
' Send data to the active (click on one in the list) client.

Private Sub cmdSend_Click()

31     On Error GoTo ErrorHandler
32 Dim lngLoop As Long

    ' You have to specify which connection you want to use
33     If lngSocket = 0 Then
34         MsgBox "First you have to select a connection!", vbCritical, "Sending data"
35         Exit Sub
36     End If

    'Send the data
37     cServer.Connection(lngSocket).Send txtSendData

    ' Go to the coresponding listview item and update it
38     For lngLoop = 1 To ListView1.ListItems.Count
39         If CLng(ListView1.ListItems(lngLoop)) = lngSocket Then
40             ListView1.ListItems(lngLoop).SubItems(6) = ListView1.ListItems(lngLoop).SubItems(6) + Len(txtSendData)
41             ListView1.ListItems(lngLoop).SubItems(7) = Now
42             Exit For
43         End If
44     Next lngLoop

45 Exit Sub

46 ErrorHandler:
47     HandleTheException "frmServer :: Error in cmdSend_Click() on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in cmdSend_Click()", enumExceptionHandling_Exception

End Sub

'Purpose:
' Send data to the all connected clients.

Private Sub cmdSendAll_Click()

48     On Error GoTo ErrorHandler
49 Dim lngLoop As Long

    'Go through all connections
50     For lngLoop = 1 To ListView1.ListItems.Count
        'Send the data
51         cServer.Connection(CLng(ListView1.ListItems(lngLoop))).Send txtSendData
        'Update the listview
52         ListView1.ListItems(lngLoop).SubItems(6) = ListView1.ListItems(lngLoop).SubItems(6) + Len(txtSendData)
53         ListView1.ListItems(lngLoop).SubItems(7) = Now
54         DoEvents
55     Next lngLoop

56 Exit Sub

57 ErrorHandler:
58     HandleTheException "frmServer :: Error in cmdSendAll_Click() on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in cmdSendAll_Click()", enumExceptionHandling_Exception

End Sub

'Purpose:
' Just stop.

Private Sub cmdDisconnectAll_Click()

59     On Error GoTo ErrorHandler

60     Unload Me

61 Exit Sub

62 ErrorHandler:
63     HandleTheException "frmServer :: Error in cmdDisconnectAll_Click() on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in cmdDisconnectAll_Click()", enumExceptionHandling_Exception

End Sub

'Purpose:
' Set the server in listening mode on the specified port.

Public Sub Listen(strPort As String)

64     On Error GoTo ErrorHandler

65     cServer.Listen CLng(strPort)
66     Me.Caption = "Server listening at port " & CLng(strPort)

67 Exit Sub

68 ErrorHandler:
69     HandleTheException "frmServer :: Error in Listen(" & strPort & ") on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in Listen(" & strPort & ")", enumExceptionHandling_Exception

End Sub

'Purpose:
' The connection where you click on will be the active connection.

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)

70     On Error GoTo ErrorHandler

71     lngSocket = CLng(Item.Text)
    'Get end point information
72     Frame1.Caption = "Selected Client " & cServer.Connection(lngSocket).GetRemoteHost & " (" & cServer.Connection(lngSocket).GetRemoteIP & ") on port " & cServer.Connection(lngSocket).GetRemotePort
73     lblLocalHost.Caption = "Local Host: " & cServer.Connection(lngSocket).GetLocalHost
74     lblLocalIP.Caption = "Local IP: " & cServer.Connection(lngSocket).GetLocalIP
75     lblLocalPort.Caption = "Local Port: " & cServer.Connection(lngSocket).GetLocalPort
76     lblRemoteHost.Caption = "Remote Host: " & cServer.Connection(lngSocket).GetRemoteHost
77     lblRemoteIP.Caption = "Remote IP: " & cServer.Connection(lngSocket).GetRemoteIP
78     lblRemotePort.Caption = "Remote Port: " & cServer.Connection(lngSocket).GetRemotePort
79     txtDataRecv = ""

80     MsgBox "The User property in the custom object (clsUserStatus) for this connection hast the value : " & cServer.Connection(lngSocket).CustomObject.User

81 Exit Sub

82 ErrorHandler:
83     HandleTheException "frmServer :: Error in ListView1_ItemClick(..) on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in ListView1_ItemClick(..)", enumExceptionHandling_Exception

End Sub

'Purpose:
' Whatever was set up as the active connection can not be active anymore.

Private Sub ClearData(lngSocketX As Long)

84     On Error GoTo ErrorHandler
85 Dim lngLoop As Long

    'Remove it from the list
86     For lngLoop = 1 To ListView1.ListItems.Count
87         If CLng(ListView1.ListItems(lngLoop)) = lngSocketX Then
88             ListView1.ListItems.Remove lngLoop
89             Exit For
90         End If
91     Next lngLoop
    'Clear the Selected client info
92     If lngSocket = lngSocketX Then
93         lngSocket = 0
94         Frame1.Caption = "Selected Client"
95         lblLocalHost.Caption = "Local Host: "
96         lblLocalIP.Caption = "Local IP: "
97         lblLocalPort.Caption = "Local Port: "
98         lblRemoteHost.Caption = "Remote Host: "
99         lblRemoteIP.Caption = "Remote IP: "
100         lblRemotePort.Caption = "Remote Port: "
101         txtDataRecv = ""
102     End If

103 Exit Sub

104 ErrorHandler:
105     HandleTheException "frmServer :: Error in ClearData(" & lngSocketX & ") on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in ClearData(" & lngSocketX & ")", enumExceptionHandling_Exception

End Sub

'----------------------------------------------------------
' The Server events
'----------------------------------------------------------

'Purpose:
' A client was closed

Private Sub cServer_OnClose(lngSocketX As Long)

106     On Error GoTo ErrorHandler

107     ClearData lngSocketX

108 Exit Sub

109 ErrorHandler:
110     HandleTheException "frmServer :: Error in cServer_OnClose(" & lngSocketX & ") on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in cServer_OnClose(" & lngSocketX & ")", enumExceptionHandling_Exception

End Sub

'Purpose:
' A client wants to connect. Accept it.

Private Sub cServer_OnConnectRequest(lngSocket As Long)

111     On Error GoTo ErrorHandler
112 Dim lngNewSocket As Long

    'Accept the connection and store the new socket handle
113     lngNewSocket = cServer.Accept(lngSocket)

    'We use the listbox to hold the info about the new client
114 Dim ListHeader    As ListItem
115     Set ListHeader = ListView1.ListItems.Add(, , lngNewSocket)
116     ListHeader.SubItems(1) = cServer.Connection(lngNewSocket).GetRemoteHost
117     ListHeader.SubItems(2) = cServer.Connection(lngNewSocket).GetRemoteIP
118     ListHeader.SubItems(3) = cServer.Connection(lngNewSocket).GetRemotePort
119     ListHeader.SubItems(4) = Now
120     ListHeader.SubItems(5) = 0
121     ListHeader.SubItems(6) = 0
122     ListHeader.SubItems(7) = Now

    'Get end point information
123     Me.Caption = "Server " & cServer.Connection(lngNewSocket).GetLocalHost & " (" & cServer.Connection(lngNewSocket).GetLocalIP & ") is listening at port " & cServer.Connection(lngNewSocket).GetLocalPort

    ' Add a new instance of the class clsUserStatus to the new connection
124 Dim UserStatusCls As New clsUserStatus
125     cServer.Connection(lngNewSocket).SetCustomObject UserStatusCls
    ' Reference a property in this connection related class
126     cServer.Connection(lngNewSocket).CustomObject.User = "Username=" & cServer.Connection(lngNewSocket).GetRemoteHost

127 Exit Sub

128 ErrorHandler:
129     HandleTheException "frmServer :: Error in cServer_OnConnectRequest(" & lngSocket & ") on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in cServer_OnConnectRequest(" & lngSocket & ")", enumExceptionHandling_Exception

End Sub

'Purpose:
' This event will be triggered when data has arived.
' This is the location where you will write your server side protocol handler.
' In this case we just log the data and update the statistics.

Private Sub cServer_OnDataArrive(lngSocketX As Long)

130     On Error GoTo ErrorHandler
131 Dim strData As String
132 Dim lngLoop As Long

    'Recieve data on the server socket
133     cServer.Connection(lngSocketX).Recv strData

    ' Go to the coresponding listview item
134     For lngLoop = 1 To ListView1.ListItems.Count
135         If CLng(ListView1.ListItems(lngLoop)) = lngSocketX Then
136             ListView1.ListItems(lngLoop).SubItems(5) = ListView1.ListItems(lngLoop).SubItems(5) + Len(strData)
137             ListView1.ListItems(lngLoop).SubItems(7) = Now
138             Exit For
139         End If
140     Next lngLoop

    ' Only show the data if it's the active/selected client
141     If lngSocket = lngSocketX Then
        'Log it
142         If Len(strData) > 0 Then
143             txtDataRecv.Text = txtDataRecv.Text & strData & vbCrLf
144             txtDataRecv.SelStart = Len(txtDataRecv.Text)
145         End If
146     End If

147 Exit Sub

148 ErrorHandler:
149     HandleTheException "frmServer :: Error in cServer_OnDataArrive(" & lngSocketX & ") on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in cServer_OnDataArrive(" & lngSocketX & ")", enumExceptionHandling_Exception

End Sub

'Purpose:
' This event is called whenever there was an error.

Private Sub cServer_OnError(lngRetCode As Long, strDescription As String)

150     On Error GoTo ErrorHandler

151     txtDataRecv.Text = txtDataRecv & "*** Error: " & strDescription & vbCrLf
152     txtDataRecv.SelStart = Len(txtDataRecv.Text)

153 Exit Sub

154 ErrorHandler:
155     HandleTheException "frmServer :: Error in cServer_OnError(..) on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in cServer_OnError(..)", enumExceptionHandling_Exception

End Sub

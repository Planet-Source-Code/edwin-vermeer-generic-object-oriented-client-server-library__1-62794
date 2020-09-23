VERSION 5.00
Begin VB.Form frmClient 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Client"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      Height          =   435
      Left            =   6360
      TabIndex        =   9
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtSendData 
      Height          =   1080
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   6255
   End
   Begin VB.TextBox txtDataRecv 
      Height          =   1575
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1200
      Width           =   7335
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   435
      Left            =   6360
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblLocalHost 
      Caption         =   "Local Host:"
      Height          =   330
      Left            =   7560
      TabIndex        =   8
      Top             =   120
      Width           =   2640
   End
   Begin VB.Label lblLocalIP 
      Caption         =   "Local IP:"
      Height          =   330
      Left            =   7560
      TabIndex        =   7
      Top             =   540
      Width           =   2640
   End
   Begin VB.Label lblLocalPort 
      Caption         =   "Local Port:"
      Height          =   330
      Left            =   7560
      TabIndex        =   6
      Top             =   960
      Width           =   2640
   End
   Begin VB.Label lblRemoteHost 
      Caption         =   "Remote Host"
      Height          =   330
      Left            =   7560
      TabIndex        =   5
      Top             =   1380
      Width           =   2640
   End
   Begin VB.Label lblRemoteIP 
      Caption         =   "Remote IP:"
      Height          =   330
      Left            =   7560
      TabIndex        =   4
      Top             =   1800
      Width           =   2640
   End
   Begin VB.Label lblRemotePort 
      Caption         =   "Remote Port:"
      Height          =   330
      Left            =   7560
      TabIndex        =   3
      Top             =   2220
      Width           =   2640
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Please read the license information and description in one of the GenericClientServer cls files
'
'Purpose:
' This demo will show how to create multiple instances of a generic client and server.
'
Option Explicit

'We will handle a client
Private WithEvents cClient  As GenericClient
Attribute cClient.VB_VarHelpID = -1

'Purpose:
' Create a new instance of the client

Private Sub Form_Load()

1     On Error GoTo ErrorHandler

2     Set cClient = New GenericClient

3 Exit Sub

4 ErrorHandler:
5     HandleTheException "frmClient :: Error in Form_Load() on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in Form_Load()", enumExceptionHandling_Exception

End Sub

'Purpose:
' Unload the client class - This MUST be done

Private Sub Form_Unload(Cancel As Integer)

6     On Error GoTo ErrorHandler

7     cClient.Connection.CloseSocket
8     Set cClient = Nothing

9 Exit Sub

10 ErrorHandler:
11     HandleTheException "frmClient :: Error in Form_Unload() on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in Form_Unload()", enumExceptionHandling_Exception

End Sub

'Purpose:
' Start the connection to the specified server

Public Sub Connect(strHost As String, strPort As String)

12     On Error GoTo ErrorHandler

13     cClient.Connect strHost, CLng(strPort)

14 Exit Sub

15 ErrorHandler:
16     HandleTheException "frmClient :: Error in Connect(""" & strHost & """, """ & strPort & """) on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in Connect(""" & strHost & """, """ & strPort & """)", enumExceptionHandling_Exception

End Sub

'Purpose:
'Send the data

Private Sub cmdSend_Click()

17     On Error GoTo ErrorHandler

18     cClient.Connection.Send txtSendData

19 Exit Sub

20 ErrorHandler:
21     HandleTheException "frmClient :: Error in Form_Unload() on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in Form_Unload()", enumExceptionHandling_Exception

End Sub

'Purpose:
'Just stop

Private Sub cmdDisconnect_Click()

22     On Error GoTo ErrorHandler

23     Unload Me

24 Exit Sub

25 ErrorHandler:
26     HandleTheException "frmClient :: Error in cmdDisconnect_Click() on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in cmdDisconnect_Click()", enumExceptionHandling_Exception

End Sub

'----------------------------------------------------------
' The Client events
'----------------------------------------------------------

'Purpose:
' This event is called by the client object when the connection was closed by the server.
' There is no connection anymore so we can quit.

Private Sub cClient_OnClose()

27     On Error GoTo ErrorHandler

28     Unload Me

29 Exit Sub

30 ErrorHandler:
31     HandleTheException "frmClient :: Error in cClient_OnClose() on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in cClient_OnClose()", enumExceptionHandling_Exception

End Sub

'Purpose:
' This event is called when the connection to the server was successfull.

Private Sub cClient_OnConnect()

32     On Error GoTo ErrorHandler
    'We are connected. Just show som information about the connection

    'Add log
33     txtDataRecv.Text = txtDataRecv.Text & "*** Connected ***" & vbCrLf
34     txtDataRecv.SelStart = Len(txtDataRecv.Text)

    'Show the end point information
35     Me.Caption = "Client connected to " & cClient.Connection.GetRemoteHost & "(" & cClient.Connection.GetRemoteIP & ")" & " on port " & cClient.Connection.GetRemotePort
36     lblLocalHost.Caption = "Local Host: " & cClient.Connection.GetLocalHost
37     lblLocalIP.Caption = "Local IP: " & cClient.Connection.GetLocalIP
38     lblLocalPort.Caption = "Local Port: " & cClient.Connection.GetLocalPort
39     lblRemoteHost.Caption = "Remote Host: " & cClient.Connection.GetRemoteHost
40     lblRemoteIP.Caption = "Remote IP: " & cClient.Connection.GetRemoteIP
41     lblRemotePort.Caption = "Remote Port: " & cClient.Connection.GetRemotePort

42 Exit Sub

43 ErrorHandler:
44     HandleTheException "frmClient :: Error in cClient_OnConnect() on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in cClient_OnConnect()", enumExceptionHandling_Exception

End Sub

'Purpose:
' this event will be called when data has arrived from the server.
' Here is where your client side protocol handler should be.
' For this simple demo whe just log the data.

Private Sub cClient_OnDataArrive()

45     On Error GoTo ErrorHandler
46 Dim strData As String

    'Recieve the data
47     cClient.Connection.Recv strData

    'Log it
48     If Len(strData) > 0 Then
49         txtDataRecv.Text = txtDataRecv.Text & strData & vbCrLf
50         txtDataRecv.SelStart = Len(txtDataRecv.Text)
51     End If

52 Exit Sub

53 ErrorHandler:
54     HandleTheException "frmClient :: Error in cClient_OnDataArrive() on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in cClient_OnDataArrive()", enumExceptionHandling_Exception

End Sub

'Purpose:
'There was an error in the client class

Private Sub cClient_OnError(lngRetCode As Long, strDescription As String)

55     On Error GoTo ErrorHandler

    'Log it
56     txtDataRecv.Text = txtDataRecv & "*** Error: " & strDescription
57     txtDataRecv.SelStart = Len(txtDataRecv.Text)

58 Exit Sub

59 ErrorHandler:
60     HandleTheException "frmClient :: Error in cClient_OnError(..) on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in cClient_OnError()", enumExceptionHandling_Exception

End Sub

VERSION 5.00
Begin VB.Form frmDemo 
   BackColor       =   &H80000018&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Client Server Demo"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Socks4 Proxy"
      Height          =   1575
      Left            =   0
      TabIndex        =   8
      Top             =   2280
      Width           =   3255
      Begin VB.CommandButton cmdListen2 
         Caption         =   "New Proxy"
         Height          =   435
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1320
         TabIndex        =   9
         Text            =   "8080"
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Local Port"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Port mapper"
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.TextBox txtRemotePort 
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Text            =   "80"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtRemoteHost 
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Text            =   "localhost"
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtLocalPort 
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Text            =   "8080"
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton cmdListen 
         Caption         =   "New Port Mapper"
         Height          =   435
         Left            =   240
         TabIndex        =   1
         Top             =   1680
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Remote Port"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Remote host"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Local Port"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Testing a proxy is easy. Just configure your client to use this machine name and the port number specified. "
      Height          =   735
      Left            =   3360
      TabIndex        =   14
      Top             =   2640
      Width           =   4455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmDemo.frx":0000
      Height          =   1455
      Left            =   3360
      TabIndex        =   13
      Top             =   1080
      Width           =   4455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmDemo.frx":01AB
      Height          =   855
      Left            =   3360
      TabIndex        =   12
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Purpose:
' This demo will show how to create multiple instances of a generic client and server.
' Created by Edwin Vermeer
' Website http://siteskinner.com
'
'Credits:
' The (super) SubClass code is from Paul Canton [Paul_Caton@hotmail.com]
' see http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=42918&lngWId=1
' Most of the winsock stuff is based on the code from 'Coding Genius'
' see http://www.planet-source-code.com/vb/scripts/showcode.asp?txtCodeId=39858&lngWId=1
' Most of the Exception hanler is from Thushan Fernando.
' see http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=41471&lngWId=1
Option Explicit

Private Sub Form_Load()

' This exception handler is even capable to capture GPF exceptions.

1     InstallExceptionHandler Me, App

End Sub

Private Sub Form_Unload(Cancel As Integer)

' You have to make sure that you uninstall this exception handler.

2     UninstallExceptionHandler

End Sub

'Purpose:
' Open a new Server form and start listning

Private Sub cmdListen_Click()

3     On Error GoTo ErrorHandler
4 Dim Server As New frmServer

5     Server.Show
6     Server.Listen txtLocalPort, txtRemoteHost, txtRemotePort, False

7 Exit Sub

8 ErrorHandler:
9     HandleTheException "frmDemo :: Error in cmdListen_Click() on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in cmdListen_Click()", enumExceptionHandling_Exception

End Sub


'Purpose:
' Open a new Server form and start listening

Private Sub cmdListen2_Click()

3     On Error GoTo ErrorHandler
4 Dim Server As New frmServer

5     Server.Show
6     Server.Listen txtLocalPort, vbNullString, vbNullString, True

7 Exit Sub

8 ErrorHandler:
9     HandleTheException "frmDemo :: Error in cmdListen_Click() on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in cmdListen_Click()", enumExceptionHandling_Exception

End Sub



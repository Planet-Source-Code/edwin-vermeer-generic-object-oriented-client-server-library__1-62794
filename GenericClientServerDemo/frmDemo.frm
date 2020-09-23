VERSION 5.00
Begin VB.Form frmDemo 
   BackColor       =   &H80000018&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Client Server Demo"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Server"
      Height          =   1620
      Left            =   0
      TabIndex        =   6
      Top             =   2040
      Width           =   3210
      Begin VB.TextBox txtLocalPort 
         Height          =   285
         Left            =   1260
         TabIndex        =   8
         Text            =   "8080"
         Top             =   420
         Width           =   1590
      End
      Begin VB.CommandButton cmdListen 
         Caption         =   "New Server"
         Height          =   435
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Local Port"
         Height          =   330
         Left            =   210
         TabIndex        =   9
         Top             =   420
         Width           =   1275
      End
   End
   Begin VB.Frame frmClient 
      Caption         =   "Client"
      Height          =   2025
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3210
      Begin VB.CommandButton cmdConnect 
         Caption         =   "New Client"
         Height          =   435
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtHostName 
         Height          =   285
         Left            =   1260
         TabIndex        =   2
         Text            =   "localhost"
         Top             =   420
         Width           =   1590
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   1260
         TabIndex        =   1
         Text            =   "8080"
         Top             =   840
         Width           =   1590
      End
      Begin VB.Label Label1 
         Caption         =   "Host Name"
         Height          =   330
         Left            =   210
         TabIndex        =   5
         Top             =   420
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "Remote Port"
         Height          =   330
         Left            =   210
         TabIndex        =   4
         Top             =   840
         Width           =   1275
      End
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmDemo.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3360
      TabIndex        =   11
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmDemo.frx":00FA
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3360
      TabIndex        =   10
      Top             =   2040
      Width           =   2895
   End
End
Attribute VB_Name = "frmDemo"
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

1     InstallExceptionHandler Me, App, True

End Sub

Private Sub Form_Unload(Cancel As Integer)

' You have to make sure that you uninstall this exception handler.

2     UninstallExceptionHandler

End Sub

'Purpose:
' Open a new Client form and start up the connection

Private Sub cmdConnect_Click()

3     On Error GoTo ErrorHandler
4 Dim Client As New frmClient

5     Client.Show
6     Client.Connect txtHostName, txtPort

7 Exit Sub

8 ErrorHandler:
9     HandleTheException "frmDemo :: Error in cmdConnect_Click() on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in cmdConnect_Click()", enumExceptionHandling_Exception

End Sub

'Purpose:
' Open a new Server form and start up the connection

Private Sub cmdListen_Click()

10     On Error GoTo ErrorHandler
11 Dim Server As New frmServer

12     Server.Show
13     Server.Listen txtLocalPort

14 Exit Sub

15 ErrorHandler:
16     HandleTheException "frmDemo :: Error in cmdListen_Click() on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in cmdListen_Click()", enumExceptionHandling_Exception

End Sub

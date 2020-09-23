VERSION 5.00
Begin VB.Form frmException 
   Caption         =   "Exception Error"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9675
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmException.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4440
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtException 
      Appearance      =   0  'Flat
      Height          =   1815
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1920
      Width           =   9495
   End
   Begin VB.CommandButton cmdContinue 
      Cancel          =   -1  'True
      Caption         =   "&Continue.."
      Height          =   330
      Left            =   8520
      TabIndex        =   5
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Default         =   -1  'True
      Height          =   330
      Left            =   7320
      TabIndex        =   4
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   35
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   9735
   End
   Begin VB.Label lblErrorTitle 
      AutoSize        =   -1  'True
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1080
      TabIndex        =   9
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Exception:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00981F0A&
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   870
   End
   Begin VB.Label lblToDo 
      BackStyle       =   0  'Transparent
      Caption         =   ".."
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   3840
      Width           =   7095
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   ".."
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   9615
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   9000
      Picture         =   "frmException.frx":0CCA
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".."
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   120
   End
   Begin VB.Label lblException 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exception ErrorHandler"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1950
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   -240
      Top             =   0
      Width           =   9975
   End
End
Attribute VB_Name = "frmException"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Purpose:
' This demo will show how to create multiple instances of a generic client and server.
' Created by Edwin Vermeer
' Website http://www.evict.nl
'
'Credits:
' Most of the Exception hanler is from Thushan Fernando.
' see http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=41471&lngWId=1

Option Explicit
Private blnContinue As Boolean

Public Property Get bContinue() As Boolean

1     bContinue = blnContinue

End Property

Private Sub cmdExit_Click()

2     blnContinue = False
3     Unload Me

End Sub

Private Sub cmdContinue_Click()

4     blnContinue = True
5     Unload Me

End Sub

Private Sub Form_Load()

6     lblTitle.Caption = "An exception occured in '" & objApp.ProductName & " " & objApp.Major & "." & objApp.Minor & "." & objApp.Revision & "'"

End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUserStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'
' Please read the license information and description in one of the GenericClientServer cls files
'
'Purpose:
' This is a verry simple class object with the only purpose to show you how an instance
' of an object can be added to every connection.
' See frmServer.cServer_OnConnectRequest for code that implements this class object
' See frmServer.ListView1_ItemClick for code that uses this class object

'local variable(s) to hold property value(s)
Option Explicit
Private mvarUser As String 'local copy

'used when assigning a value to the property, on the left side of an assignment.

Public Property Let User(ByVal vData As String)

1     mvarUser = vData

End Property

'used when retrieving value of a property, on the right side of an assignment.

Public Property Get User() As String

2     User = mvarUser

End Property

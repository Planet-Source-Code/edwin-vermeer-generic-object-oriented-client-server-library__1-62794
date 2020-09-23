VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExceptionHandler"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Purpose:
' This demo will show how to create multiple instances of a generic client and server.
' Created by Edwin Vermeer
' Website http://www.evict.nl
'
'Credits:
' Most of the Exception hanler is from Thushan Fernando.
' see http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=41471&lngWId=1

Option Explicit

Public Enum enumExceptionHandling
    enumExceptionHandling_Exception = 1
    enumExceptionHandling_Warning = 2
    enumExceptionHandling_ScriptError = 3
    enumExceptionHandling_Information = 4
End Enum

Public Enum enumExceptionType
    enumExceptionType_AccessViolation = &HC0000005
    enumExceptionType_DataTypeMisalignment = &H80000002
    enumExceptionType_Breakpoint = &H80000003
    enumExceptionType_SingleStep = &H80000004
    enumExceptionType_ArrayBoundsExceeded = &HC000008C
    enumExceptionType_FaultDenormalOperand = &HC000008D
    enumExceptionType_FaultDivideByZero = &HC000008E
    enumExceptionType_FaultInexactResult = &HC000008F
    enumExceptionType_FaultInvalidOperation = &HC0000090
    enumExceptionType_FaultOverflow = &HC0000091
    enumExceptionType_FaultStackCheck = &HC0000092
    enumExceptionType_FaultUnderflow = &HC0000093
    enumExceptionType_IntegerDivisionByZero = &HC0000094
    enumExceptionType_IntegerOverflow = &HC0000095
    enumExceptionType_PriviledgedInstruction = &HC0000096
    enumExceptionType_InPageError = &HC0000006
    enumExceptionType_IllegalInstruction = &HC000001D
    enumExceptionType_NoncontinuableException = &HC0000025
    enumExceptionType_StackOverflow = &HC00000FD
    enumExceptionType_InvalidDisposition = &HC0000026
    enumExceptionType_GuardPageViolation = &H80000001
    enumExceptionType_InvalidHandle = &HC0000008
    enumExceptionType_ControlCExit = &HC000013A
End Enum

Public Sub InstallExceptionHandler(frmForm As Object, objApplication As Object, blnLogInEventlog As Boolean)
Attribute InstallExceptionHandler.VB_Description = "In order to be able to capture GPF errors you must install the ExceptionHandler. "

1     Set frmMain = frmForm
2     Set objApp = objApplication
3     blnLogEvents = blnLogInEventlog
4     modInstallExceptionHandler

End Sub

Public Sub UninstallExceptionHandler()
Attribute UninstallExceptionHandler.VB_Description = "When the exception handler is installed then we also must uninstall it."

5     modUninstallExceptionHandler

End Sub

Public Sub HandleTheException(strException As String, strProcedure As String, intType As enumExceptionHandling)
Attribute HandleTheException.VB_Description = "We will call this method to handle an exception."

6     If blnLogEvents = True Then
7         Select Case intType
        Case enumExceptionHandling_Warning
8             WriteLog EventLog_Application, objApp.Title, vbLogEventTypeWarning, strException
9         Case enumExceptionHandling_Exception, enumExceptionHandling_ScriptError
10             WriteLog EventLog_Application, objApp.Title, vbLogEventTypeError, strException
11         Case enumExceptionHandling_Information
12             WriteLog EventLog_Application, objApp.Title, vbLogEventTypeInformation, strException
13         End Select
14     End If

15     modHandleTheException strException, strProcedure, intType

End Sub

Public Function GetExceptionName(ExceptionType As enumExceptionType) As String
Attribute GetExceptionName.VB_Description = "When there was a GPF exception. You can get the name of the exception."

16     GetExceptionName = modGetExceptionName(ExceptionType)

End Function

Public Sub RaiseAnException(ExceptionType As enumExceptionType)
Attribute RaiseAnException.VB_Description = "This is usually only used for testing. You can simulate an Exception."

17     modRaiseAnException ExceptionType

End Sub

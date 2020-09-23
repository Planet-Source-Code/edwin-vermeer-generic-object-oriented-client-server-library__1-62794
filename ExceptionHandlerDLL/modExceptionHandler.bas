Attribute VB_Name = "modExceptionHandler"
'Purpose:
' This demo will show how to create multiple instances of a generic client and server.
' Created by Edwin Vermeer
' Website http://www.evict.nl
'
'Credits:
' Most of the Exception hanler is from Thushan Fernando.
' see http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=41471&lngWId=1

Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As EXCEPTION_RECORD, ByVal LPEXCEPTION_RECORD As Long, ByVal lngBytes As Long)
Private Declare Function SetUnhandledExceptionFilter Lib "kernel32" (ByVal lpTopLevelExceptionFilter As Long) As Long
Private Declare Sub RaiseException Lib "kernel32" (ByVal dwExceptionCode As Long, ByVal dwExceptionFlags As Long, ByVal nNumberOfArguments As Long, lpArguments As Long)

Public Enum enumExceptionTypeAlt
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

Public Enum enumExceptionHandlingAlt
    enumExceptionHandling_Warning = 1
    enumExceptionHandling_Exception = 2
    enumExceptionHandling_ScriptError = 3
End Enum

Private Const EXCEPTION_MAXIMUM_PARAMETERS = 15

Private Type CONTEXT
    FltF0        As Double
    FltF1        As Double
    FltF2        As Double
    FltF3        As Double
    FltF4        As Double
    FltF5        As Double
    FltF6        As Double
    FltF7        As Double
    FltF8        As Double
    FltF9        As Double
    FltF10       As Double
    FltF11       As Double
    FltF12       As Double
    FltF13       As Double
    FltF14       As Double
    FltF15       As Double
    FltF16       As Double
    FltF17       As Double
    FltF18       As Double
    FltF19       As Double
    FltF20       As Double
    FltF21       As Double
    FltF22       As Double
    FltF23       As Double
    FltF24       As Double
    FltF25       As Double
    FltF26       As Double
    FltF27       As Double
    FltF28       As Double
    FltF29       As Double
    FltF30       As Double
    FltF31       As Double
    IntV0        As Double
    IntT0        As Double
    IntT1        As Double
    IntT2        As Double
    IntT3        As Double
    IntT4        As Double
    IntT5        As Double
    IntT6        As Double
    IntT7        As Double
    IntS0        As Double
    IntS1        As Double
    IntS2        As Double
    IntS3        As Double
    IntS4        As Double
    IntS5        As Double
    IntFp        As Double
    IntA0        As Double
    IntA1        As Double
    IntA2        As Double
    IntA3        As Double
    IntA4        As Double
    IntA5        As Double
    IntT8        As Double
    IntT9        As Double
    IntT10       As Double
    IntT11       As Double
    IntRa        As Double
    IntT12       As Double
    IntAt        As Double
    IntGp        As Double
    IntSp        As Double
    IntZero      As Double
    Fpcr         As Double
    SoftFpcr     As Double
    Fir          As Double
    Psr          As Long
    ContextFlags As Long
    Fill(4)      As Long
End Type
Private Type EXCEPTION_RECORD
    ExceptionCode                                        As Long
    ExceptionFlags                                       As Long
    pExceptionRecord                                     As Long
    ExceptionAddress                                     As Long
    NumberParameters                                     As Long
    ExceptionInformation(EXCEPTION_MAXIMUM_PARAMETERS)   As Long
End Type

Private Type EXCEPTION_POINTERS
    pExceptionRecord     As EXCEPTION_RECORD
    ContextRecord        As CONTEXT
End Type

Public blnIsHandlerInstalled As Boolean
Public frmMain As Form
Public objApp As Object
Public blnLogEvents As Boolean

Public Sub modHandleTheException(strException As String, strProcedure As String, intType As enumExceptionHandlingAlt)

1 Dim strEx As String

2     With frmException
3         If intType = enumExceptionHandling_Warning Then
4             .lblWarning = objApp.CompanyName & " apologizes for the inconvience but we have to warn that:"
5             .lblToDo = "Please follow the instructions in the error message, if this problem reoccurs then please contact " & objApp.CompanyName & " with a detailed description of your system and how to reproduce the issue."
6             .cmdExit.Visible = False
7         Else
8             If intType = enumExceptionHandling_Exception Then
9                 .lblWarning = objApp.CompanyName & " apologizes for the inconvience but an Error occured in the application. It may be possible for you to continue to run this application without issues but we recommend you only do so if your certain that its OK to do so. Click 'Continue' to resume ignoring the error or 'Exit' to terminate the application immediately."
10                 .lblToDo = "If the error reoccurs, then please contact " & objApp.CompanyName & " with a detailed description of your system and how to reproduce the issue."
11                 .cmdExit.Visible = True
12             Else
13                 .lblWarning = objApp.CompanyName & " apologizes for the inconvience but an Exception Error occured in the script that you are trying to execute."
14                 .lblToDo = "Please update the script and try running it again. If the error continues, then please contact the developer of the script with a detailed description how to reproduce the issue."
15                 .cmdExit.Visible = True
16             End If
17         End If
18         strEx = Replace(strException, " - Error", vbCrLf & "Error")
19         .lblErrorTitle = strProcedure & " occured on " & Date & " " & Time
20         .txtException.Text = "####  Error in " & objApp.Title & " " & objApp.Major & "." & objApp.Minor & "." & objApp.Revision & " occured on " & Date & " " & Time & "  ####" & vbCrLf & strEx
21         .Show vbModal
22         If Not .bContinue Then
23             Unload frmMain
24             DoEvents
            ' End ' Can not be used in a dll :(
25         End If
26     End With

End Sub

Public Function ExceptionHandler(ByRef ExceptionPtrs As EXCEPTION_POINTERS) As Long

27     On Error Resume Next
28     Dim ExceptionRecord As EXCEPTION_RECORD
29     Dim strExceptionDescriptiosn As String
30         ExceptionRecord = ExceptionPtrs.pExceptionRecord
31         Do Until ExceptionRecord.pExceptionRecord = 0
32             CopyMemory ExceptionRecord, ExceptionRecord.pExceptionRecord, Len(ExceptionRecord)
33         Loop
34         strExceptionDescriptiosn = GetExceptionDescription(ExceptionRecord.ExceptionCode)
35     On Error GoTo 0
36     Err.Raise vbObjectError, "ExceptionHandler", "Exception: " & strExceptionDescriptiosn & " [" & modGetExceptionName(ExceptionRecord.ExceptionCode) & "]" & vbCrLf & "ExceptionAddress : " & ExceptionRecord.ExceptionAddress

End Function

Public Sub modInstallExceptionHandler()

37     On Error Resume Next
38         If Not blnIsHandlerInstalled Then
39             Call SetUnhandledExceptionFilter(AddressOf ExceptionHandler)
40             blnIsHandlerInstalled = True
41         End If

End Sub

Public Function GetExceptionDescription(ExceptionType As enumExceptionType) As String

42     On Error Resume Next
43     Dim strDescription As String
44         Select Case ExceptionType
        Case enumExceptionType_AccessViolation
45             strDescription = "Access Violation"
46         Case enumExceptionType_DataTypeMisalignment
47             strDescription = "Data Type Misalignment"
48         Case enumExceptionType_Breakpoint
49             strDescription = "Breakpoint"
50         Case enumExceptionType_SingleStep
51             strDescription = "Single Step"
52         Case enumExceptionType_ArrayBoundsExceeded
53             strDescription = "Array Bounds Exceeded"
54         Case enumExceptionType_FaultDenormalOperand
55             strDescription = "Float Denormal Operand"
56         Case enumExceptionType_FaultDivideByZero
57             strDescription = "Divide By Zero"
58         Case enumExceptionType_FaultInexactResult
59             strDescription = "Floating Point Inexact Result"
60         Case enumExceptionType_FaultInvalidOperation
61             strDescription = "Invalid Operation"
62         Case enumExceptionType_FaultOverflow
63             strDescription = "Float Overflow"
64         Case enumExceptionType_FaultStackCheck
65             strDescription = "Float Stack Check"
66         Case enumExceptionType_FaultUnderflow
67             strDescription = "Float Underflow"
68         Case enumExceptionType_IntegerDivisionByZero
69             strDescription = "Integer Divide By Zero"
70         Case enumExceptionType_IntegerOverflow
71             strDescription = "Integer Overflow"
72         Case enumExceptionType_PriviledgedInstruction
73             strDescription = "Privileged Instruction"
74         Case enumExceptionType_InPageError
75             strDescription = "In Page Error"
76         Case enumExceptionType_IllegalInstruction
77             strDescription = "Illegal Instruction"
78         Case enumExceptionType_NoncontinuableException
79             strDescription = "Non Continuable Exception"
80         Case enumExceptionType_StackOverflow
81             strDescription = "Stack Overflow"
82         Case enumExceptionType_InvalidDisposition
83             strDescription = "Invalid Disposition"
84         Case enumExceptionType_GuardPageViolation
85             strDescription = "Guard Page Violation"
86         Case enumExceptionType_InvalidHandle
87             strDescription = "Invalid Handle"
88         Case enumExceptionType_ControlCExit
89             strDescription = "Control-C Exit"
90         Case Else
91             strDescription = "Unknown Exception Error"
92         End Select
93         GetExceptionDescription = strDescription

End Function

Public Function modGetExceptionName(ExceptionType As enumExceptionType) As String

94     On Error Resume Next
95     Dim strDescription As String
96         Select Case ExceptionType
        Case enumExceptionType_AccessViolation
97             strDescription = "EXCEPTION_ACCESS_VIOLATION"
98         Case enumExceptionType_DataTypeMisalignment
99             strDescription = "EXCEPTION_DATATYPE_MISALIGNMENT"
100         Case enumExceptionType_Breakpoint
101             strDescription = "EXCEPTION_BREAKPOINT"
102         Case enumExceptionType_SingleStep
103             strDescription = "EXCEPTION_SINGLE_STEP"
104         Case enumExceptionType_ArrayBoundsExceeded
105             strDescription = "EXCEPTION_ARRAY_BOUNDS_EXCEEDED"
106         Case enumExceptionType_FaultDenormalOperand
107             strDescription = "EXCEPTION_FLT_DENORMAL_OPERAND"
108         Case enumExceptionType_FaultDivideByZero
109             strDescription = "EXCEPTION_FLT_DIVIDE_BY_ZERO"
110         Case enumExceptionType_FaultInexactResult
111             strDescription = "EXCEPTION_FLT_INEXACT_RESULT"
112         Case enumExceptionType_FaultInvalidOperation
113             strDescription = "EXCEPTION_FLT_INVALID_OPERATION"
114         Case enumExceptionType_FaultOverflow
115             strDescription = "EXCEPTION_FLT_OVERFLOW"
116         Case enumExceptionType_FaultStackCheck
117             strDescription = "EXCEPTION_FLT_STACK_CHECK"
118         Case enumExceptionType_FaultUnderflow
119             strDescription = "EXCEPTION_FLT_UNDERFLOW"
120         Case enumExceptionType_IntegerDivisionByZero
121             strDescription = "EXCEPTION_INT_DIVIDE_BY_ZERO"
122         Case enumExceptionType_IntegerOverflow
123             strDescription = "EXCEPTION_INT_OVERFLOW"
124         Case enumExceptionType_PriviledgedInstruction
125             strDescription = "EXCEPTION_PRIVILEGED_INSTRUCTION"
126         Case enumExceptionType_InPageError
127             strDescription = "EXCEPTION_IN_PAGE_ERROR"
128         Case enumExceptionType_IllegalInstruction
129             strDescription = "EXCEPTION_ILLEGAL_INSTRUCTION"
130         Case enumExceptionType_NoncontinuableException
131             strDescription = "EXCEPTION_NONCONTINUABLE_EXCEPTION"
132         Case enumExceptionType_StackOverflow
133             strDescription = "EXCEPTION_STACK_OVERFLOW"
134         Case enumExceptionType_InvalidDisposition
135             strDescription = "EXCEPTION_INVALID_DISPOSITION"
136         Case enumExceptionType_GuardPageViolation
137             strDescription = "EXCEPTION_GUARD_PAGE_VIOLATION"
138         Case enumExceptionType_InvalidHandle
139             strDescription = "EXCEPTION_INVALID_HANDLE"
140         Case enumExceptionType_ControlCExit
141             strDescription = "EXCEPTION_CONTROL_C_EXIT"
142         Case Else
143             strDescription = "Unknown"
144         End Select
145         modGetExceptionName = strDescription

End Function

Public Sub modRaiseAnException(ExceptionType As enumExceptionType)

146     RaiseException ExceptionType, 0, 0, 0

End Sub

Public Sub modUninstallExceptionHandler()

147     On Error Resume Next
148         If blnIsHandlerInstalled Then
149             Call SetUnhandledExceptionFilter(0&)
150             blnIsHandlerInstalled = False
151         End If

End Sub

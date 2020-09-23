VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Subclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------------
'cSubclass - module-less, IDE safe, machine code window subclasser
'
'v1.00 20030107 First cut..........................................................................
'
'Note : This Subclass class is an unmodified copy of Paul Caton's subclasser
Option Explicit

Private Const OFFSET_P1   As Long = 9                     'Callback gate address
Private Const OFFSET_P2   As Long = 22                    'Before table entry count
Private Const OFFSET_P3   As Long = 37                    'Before table address
Private Const OFFSET_P4   As Long = 51                    'In IDE ?
Private Const OFFSET_P5   As Long = 69                    'Owner object address for iSubclass_Before
Private Const OFFSET_P6   As Long = 141                   'Original WndProc address
Private Const OFFSET_P7   As Long = 146                   'CallWindowProc address
Private Const OFFSET_P8   As Long = 154                   'After table entry count
Private Const OFFSET_P9   As Long = 169                   'After table address
Private Const OFFSET_PA   As Long = 183                   'In IDE?
Private Const OFFSET_PB   As Long = 201                   'Owner object address for iSubclass_After
Private Const OFFSET_PC   As Long = 250                   'Original WndProc address
Private Const OFFSET_PD   As Long = 260                   'SetWindowLong address
Private Const ARRAY_LB    As Long = 1                     'Low bound of arrays

Private Type tCode
    buf(ARRAY_LB To 272)    As Byte                         'Code buffer
End Type

Private Type tCodeBuf
    Code                    As tCode                        'Subclass WndProc code
End Type

Private CodeBuf           As tCodeBuf                     'Subclass WndProc code instance
Private nBreakGate        As Long                         'Callback breakpoint gate
Private nMsgCntB          As Long                         'Before msg table entry count
Private nMsgCntA          As Long                         'After msg table entry count
Private aMsgTblB()        As WinSubHook.eMsg              'Before msg table array
Private aMsgTblA()        As WinSubHook.eMsg              'After msg table array
Private hWndSubclass      As Long                         'Handle of the window being subclassed
Private nWndProcSubclass  As Long                         'The address of our WndProc
Private nWndProcOriginal  As Long                         'The address of the existing WndProc

'-----------------------------
'Class creation/destruction

'Called automatically when the class instance is created.

Private Sub Class_Initialize()

1     On Error GoTo ErrorHandler

2 Const OPS As String = "558BEC83C4F85756BE_patch1_33C08945FC8945F8B90000000083F900746183F9FF740CBF000000008B450CF2AF755033C03D_patch4_740B833E007542C70601000000BA_patch5_8B0283F8000F84A50000008D4514508D4510508D450C508D4508508D45FC508D45F8508B0252FF5020C706000000008B45F883F8007570FF7514FF7510FF750CFF750868_patch6_E8_patch7_8945FCB90000000083F900744D83F9FF740CBF000000008B450CF2AF753C33C03D_patchA_740B833E00752EC70601000000BA_patchB_8B0283F8007425FF7514FF7510FF750CFF75088D45FC508B0252FF501CC706000000005E5F8B45FCC9C2100068_patchC_6AFCFF7508E8_patchD_33C08945FCEBE190"
3 Dim i     As Long, _
    j     As Long, _
    nIDE  As Long

    'Convert the string from hexadecimal pairs to bytes and store in the opcode buffer
4     With CodeBuf.Code
5         j = 1                                                 'Set the character index to the start of the opcode string
6         For i = ARRAY_LB To UBound(.buf)                      'For each byte of the code buffer
7             .buf(i) = Val("&H" & Mid$(OPS, j, 2))               'Pull a pair of hex characters and convert to a byte
8             j = j + 2                                           'Bump to the next pair of characters
9         Next i                                                'Next byte of the code buffer
10         nWndProcSubclass = VarPtr(.buf(ARRAY_LB))             'Address of the cSubclass WndProc entry point
11     End With

12     nIDE = InIDE                                            'Determine whether we're running in the IDE or not

    'Patch the WndProc thunk with runtime values
13     Call PatchVal(OFFSET_P1, VarPtr(nBreakGate))            'Breakpoint gate address
14     Call PatchVal(OFFSET_P4, nIDE)                          'Wether we need check the breakpoint gate and the vtable
15     Call PatchRel(OFFSET_P7, AddrFunc("CallWindowProcA"))   'Address of the CallWindowProc api function
16     Call PatchVal(OFFSET_PA, nIDE)                          'Whether we need check the breakpoint gate and the vtable
17     Call PatchRel(OFFSET_PD, AddrFunc("SetWindowLongA"))    'Address of the SetWindowLong api function

18 Exit Sub

19 ErrorHandler:
20     Err.Raise vbObjectError Or Err, "SubClass", "SubClass :: Error in Class_Initialize on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Sub

'Called automatically when the class instance is destroyed.

Private Sub Class_Terminate()

21     On Error GoTo ErrorHandler
22     If hWndSubclass <> 0 Then                               'If the Subclass thunk is active
23         Call UnSubclass                                       'UnSubclass
24     End If

25 Exit Sub

26 ErrorHandler:
27     Err.Raise vbObjectError Or Err, "SubClass", "SubClass :: Error in Class_Terminate on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Sub

'-----------------------------
'Public interface

'Call this method to add a message to the msg callback table. NB This method can be called at any time

Public Sub AddMsg(uMsg As WinSubHook.eMsg, When As WinSubHook.eMsgWhen)

28     On Error GoTo ErrorHandler
29     If When = WinSubHook.MSG_BEFORE Then                    'If before
30         Call AddMsgSub(uMsg, aMsgTblB, nMsgCntB, When)        'Add the message, pass the before table and before table message count variables
31     Else                                                    'Else after
32         Call AddMsgSub(uMsg, aMsgTblA, nMsgCntA, When)        'Add the message, pass the after table and after table message count variables
33     End If

34 Exit Sub

35 ErrorHandler:
36     Err.Raise vbObjectError Or Err, "SubClass", "SubClass :: Error in AddMsg on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Sub

'Allow the user to arbitarily call the original WndProc

Public Function CallOrigWndProc(ByVal uMsg As WinSubHook.eMsg, ByVal wParam As Long, ByVal lParam As Long) As Long

37     On Error GoTo ErrorHandler

38     If hWndSubclass <> 0 Then
39         CallOrigWndProc = WinSubHook.CallWindowProc( _
                          nWndProcOriginal, hWndSubclass, uMsg, wParam, lParam) 'Call the original WndProc
40     Else
41         Debug.Assert False                                    'LOGIC ERROR: The subclasser isn't active!
42     End If

43 Exit Function

44 ErrorHandler:
45     Err.Raise vbObjectError Or Err, "SubClass", "SubClass :: Error in CallOrigWndProc on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Function

'Call this method to delete a message from the msg table. NB This method can be called at any time

Public Sub DelMsg(uMsg As WinSubHook.eMsg, When As WinSubHook.eMsgWhen)

46     On Error GoTo ErrorHandler
47     If When = WinSubHook.MSG_BEFORE Then                    'If before
48         Call DelMsgSub(uMsg, aMsgTblB, nMsgCntB, When)        'Delete the message, pass the before table and before message count variables
49     Else                                                    'Else after
50         Call DelMsgSub(uMsg, aMsgTblA, nMsgCntA, When)        'Delete the message, pass the after table and after message count variables
51     End If

52 Exit Sub

53 ErrorHandler:
54     Err.Raise vbObjectError Or Err, "SubClass", "SubClass :: Error in DelMsg on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Sub

'Call this method to subclass the passed window handle

Public Sub Subclass(hWnd As Long, Owner As WinSubHook.iSubclass)

55     On Error GoTo ErrorHandler
56     Debug.Assert (hWndSubclass = 0)                         'LOGIC ERROR: The subclasser is already active!
57     Debug.Assert IsWindow(hWnd)                             'LOGIC ERROR: Invalid window handle

58     hWndSubclass = hWnd                                     'Store the window handle
59     nWndProcOriginal = WinSubHook.SetWindowLong( _
                       hWnd, _
                       WinSubHook.GWL_WNDPROC, _
                       nWndProcSubclass)       'Set our WndProc in place of the original
60     Debug.Assert nWndProcOriginal                           '??? You can't subclass a window outside of the current process

61     Call PatchVal(OFFSET_P5, ObjPtr(Owner))                 'Owner object address for iSubclass_Before
62     Call PatchVal(OFFSET_P6, nWndProcOriginal)              'Original WndProc address for CallWindowProc
63     Call PatchVal(OFFSET_PB, ObjPtr(Owner))                 'Owner object address for iSubclass_After
64     Call PatchVal(OFFSET_PC, nWndProcOriginal)              'Original WndProc address for SetWindowLong

65 Exit Sub

66 ErrorHandler:
67     Err.Raise vbObjectError Or Err, "SubClass", "SubClass :: Error in SubClass on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Sub

'Call this method to stop subclassing the window

Public Sub UnSubclass()

68     On Error GoTo ErrorHandler
69     If hWndSubclass <> 0 Then
70         Call PatchVal(OFFSET_P2, 0)                           'Patch the code to ensure no further iSubclass_Before callbacks
71         Call PatchVal(OFFSET_P8, 0)                           'Patch the code to ensure no further iSubclass_After callbacks
72         Call WinSubHook.SetWindowLong(hWndSubclass, _
                                      WinSubHook.GWL_WNDPROC, _
                                      nWndProcOriginal)       'Restore the original WndProc
73         hWndSubclass = 0                                      'Indicate the subclasser is inactive
74         nMsgCntB = 0                                          'Message before count equals zero
75         nMsgCntA = 0                                          'Message after count equals zero
76     End If

77 Exit Sub

78 ErrorHandler:
79     Err.Raise vbObjectError Or Err, "SubClass", "SubClass :: Error in UnSubclass on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Sub

'-----------------------------
' Private subroutines

'Worker sub for AddMsg

Private Sub AddMsgSub(uMsg As WinSubHook.eMsg, aMsgTbl() As WinSubHook.eMsg, nMsgCnt As Long, When As WinSubHook.eMsgWhen)

80     On Error GoTo ErrorHandler
81 Dim nEntry  As Long, _
    nOff1   As Long, _
    nOff2   As Long

82     If uMsg = WinSubHook.ALL_MESSAGES Then                  'If ALL_MESSAGES
83         nMsgCnt = -1                                          'Indicates that all messages are to callback
84     Else                                                    'Else a specific message number
85         For nEntry = ARRAY_LB To nMsgCnt                      'For each existing entry. NB will skip if nMsgCnt = 0
86             Select Case aMsgTbl(nEntry)                         'Select on the message number stored in this table entry
            Case -1                                             'This msg table slot is a deleted entry
87                 aMsgTbl(nEntry) = uMsg                            'Re-use this entry
88                 Exit Sub                                          'Bail
89             Case uMsg                                           'The msg is already in the table!
90                 Exit Sub                                          'Bail
91             End Select
92         Next nEntry                                           'Next entry

        'Make space for the new entry
93         ReDim Preserve aMsgTbl(ARRAY_LB To nEntry)            'Increase the size of the table. NB nEntry = nMsgCnt + 1
94         nMsgCnt = nEntry                                      'Bump the entry count
95         aMsgTbl(nEntry) = uMsg                                'Store the message number in the table
96     End If

97     If When = WinSubHook.MSG_BEFORE Then                    'If before
98         nOff1 = OFFSET_P2                                     'Patch the Before table entry count
99         nOff2 = OFFSET_P3                                     'Patch the Before table address
100     Else                                                    'Else after
101         nOff1 = OFFSET_P8                                     'Patch the After table entry count
102         nOff2 = OFFSET_P9                                     'Patch the After table address
103     End If

    'Patch the appropriate table entries
104     Call PatchVal(nOff1, nMsgCnt)                           'Patch the appropriate table entry count
105     Call PatchVal(nOff2, AddrMsgTbl(aMsgTbl))               'Patch the appropriate table address. We need do this because there's no guarantee that the table existed at SubClass time, the table only gets created if a specific message number is added.

106 Exit Sub

107 ErrorHandler:
108     Err.Raise vbObjectError Or Err, "SubClass", "SubClass :: Error in AddMsgSub on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Sub

'Worker sub for DelMsg

Private Sub DelMsgSub(uMsg As WinSubHook.eMsg, aMsgTbl() As WinSubHook.eMsg, nMsgCnt As Long, When As WinSubHook.eMsgWhen)

109     On Error GoTo ErrorHandler
110 Dim nEntry As Long

111     If uMsg = WinSubHook.ALL_MESSAGES Then                  'If deleting all messages (specific or ALL_MESSAGES)
112         nMsgCnt = 0                                           'Message count is now zero
113         If When = WinSubHook.MSG_BEFORE Then                  'If before
114             nEntry = OFFSET_P2                                  'Patch the before table message count location
115         Else                                                  'Else after
116             nEntry = OFFSET_P8                                  'Patch the after table message count location
117         End If
118         Call PatchVal(nEntry, 0)                              'Patch the table message count
119     Else                                                    'Else deleteting a specific message
120         For nEntry = ARRAY_LB To nMsgCnt                      'For each table entry
121             If aMsgTbl(nEntry) = uMsg Then                      'If this entry is the message we wish to delete
122                 aMsgTbl(nEntry) = -1                              'Mark the table slot as available
123                 Exit For                                          'Bail
124             End If
125         Next nEntry                                           'Next entry
126     End If

127 Exit Sub

128 ErrorHandler:
129     Err.Raise vbObjectError Or Err, "SubClass", "SubClass :: Error in DelMsgSub on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Sub

'Return the address of the passed user32.dll api function

Private Function AddrFunc(sProc As String) As Long

130     On Error GoTo ErrorHandler

131     AddrFunc = WinSubHook.GetProcAddress(WinSubHook.GetModuleHandle("user32"), sProc)

132 Exit Function

133 ErrorHandler:
134     Err.Raise vbObjectError Or Err, "SubClass", "SubClass :: Error in AddrFunc on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Function

'Return the address of the low bound of the passed table array

Private Function AddrMsgTbl(aMsgTbl() As WinSubHook.eMsg) As Long

135     On Error Resume Next                                    'The table may not be dimensioned yet so we need protection
136         AddrMsgTbl = VarPtr(aMsgTbl(ARRAY_LB))                'Get the address of the first element of the passed message table
137     On Error GoTo 0                                         'Switch off error protection

End Function

'Patch the code offset with the passed value

Private Sub PatchVal(nOffset As Long, nValue As Long)

138     On Error GoTo ErrorHandler

139     Call WinSubHook.CopyMemory(ByVal (nWndProcSubclass + nOffset), nValue, 4)

140 Exit Sub

141 ErrorHandler:
142     Err.Raise vbObjectError Or Err, "SubClass", "SubClass :: Error in PatchVal on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Sub

'Patch the code offset with the relative address to the target address

Private Sub PatchRel(nOffset As Long, nTargetAddr As Long)

143     On Error GoTo ErrorHandler

144     Call WinSubHook.CopyMemory(ByVal (nWndProcSubclass + nOffset), nTargetAddr - nWndProcSubclass - nOffset - 4, 4)

145 Exit Sub

146 ErrorHandler:
147     Err.Raise vbObjectError Or Err, "SubClass", "SubClass :: Error in PatchRel on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Sub

'Return -1 if we're running in the IDE or 0 if were running compiled.

Private Function InIDE() As Long

148     On Error GoTo ErrorHandler
149 Static Value As Long

150     If Value = 0 Then
151         Value = 1
152         Debug.Assert True Or InIDE()                          'This line won't exist in the compiled app
153         InIDE = Value - 1
154     End If

155     Value = 0

156 Exit Function

157 ErrorHandler:
158     Err.Raise vbObjectError Or Err, "SubClass", "SubClass :: Error in InIDE on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description

End Function

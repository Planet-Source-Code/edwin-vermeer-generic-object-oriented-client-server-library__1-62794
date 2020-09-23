Attribute VB_Name = "test"
Option Explicit

Public Sub main()

1     WriteLog EventLog_Application, "My Special APP", vbLogEventTypeError, "Oep, Something went wrong :)"
2     MsgBox "Just look in the eventlog to see what has been written.", vbInformation, "That's all folks."

End Sub

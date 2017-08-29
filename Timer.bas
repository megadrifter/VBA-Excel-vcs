Attribute VB_Name = "Timer"

Private Declare Function GetTickCount Lib "kernel32" () As Long


Function millis()

Debug.Print (Now)
millis = GetTickCount
Debug.Print millis
End Function


Private Sub ghdghgdfhj()
Debug.Print millis()
End Sub

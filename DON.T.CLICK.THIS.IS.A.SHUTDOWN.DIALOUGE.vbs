Dim X
X = MsgBox("Error while opening file. Do you want to dance?", vbYesNo + vbQuestion, "Computer Message")

If X = vbYes Then
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run "cmd /k curl ASCII.live/can-you-hear-me", 1, True
End If

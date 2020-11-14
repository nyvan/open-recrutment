Attribute VB_Name = "Metu"
Option Explicit
Dim x As String

Sub Keluar()
x = MsgBoxGT("Do you really want to exit?", vbQuestion + vbYesNo, "Exit Application")
If x = vbYes Then
End
End If
If x = vbNo Then
FORT.Show
End If
End Sub



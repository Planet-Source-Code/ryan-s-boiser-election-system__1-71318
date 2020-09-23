Attribute VB_Name = "Module1"
Dim isadmin As Boolean
Public Sub exitQ()
Dim askme As Single
    askme = MsgBox("Are you sure you want to close application?", vbOKCancel + vbQuestion, "SSITE")
        If askme = vbOK Then
            End
        End If
End Sub


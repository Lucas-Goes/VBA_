'Impedir a utilização do botão X (Fechar) 
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Application.DisplayAlerts = False
Application.ScreenUpdating = False
        If CloseMode = 0 Then
        Cancel = True
        MsgBox "frase para auxílio", vbExclamation
    End If
     
End Sub
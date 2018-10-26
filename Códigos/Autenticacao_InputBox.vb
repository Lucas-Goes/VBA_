Sub Autenticacao_InputBox()


Objeto = InputBox _
("Acesso restrito !" & Chr(13) & "entre em contado com suporte@suporte.com" & Chr(13) & "Senha:", "Atenção!")
' Chr(13), caracter para pular uma linha

If Objeto <> "Sua senha" Then
    MsgBox "Senha Inválida!", vbCritical + vbOKOnly
Else
    'Código pós validação
End If

End Sub
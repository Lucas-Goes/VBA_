Sub Tipos_Ocultar()

Set wsCadastro = ThisWorkbook.Sheets("Folha2")
wsCadastro.Visible = xlSheetInvisible
wsCadastro.Visible = xlSheetVisible
wsCadastro.Visible = xlVeryHidden
wsCadastro.Visible = xlSheetVisible

'xlSheetInvisible, Oculta a Folha e mantém no menu Mostrar
'xlSheetVisible, Reexibe a Folha
'xlVeryHidden, Oculta definitivamente, não é possivel reexibir através do menu Mostrar

End Sub

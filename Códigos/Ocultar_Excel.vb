'Código para ocultar o excel e exibir somente os formulários criados
Private Sub Workbook_Open()
If Workbooks.Count > 1 Then 'se houver mais de uma pasta de trabalho aberta ...
Windows(ThisWorkbook.Name).visible = False 'oculta somente esse arquivo
Example_Form.Show 'Exemplo de chamada de formulário pós validação
        Else   'somente 1 wb aberto
            Application.visible = False 'oculta
            Example_Form.Show 'Exemplo de chamada de formulário pós validação
            End If
End Sub

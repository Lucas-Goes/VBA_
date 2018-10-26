Sub Arrasta_Dados()

wb = ThisWorkbook.Name
'Arrasta as formulas até a última linha
    linhavazia = 1 'primeira linha a verificar
    Do While Not IsEmpty(Range("A" & linhavazia))
        linhavazia = linhavazia + 1
    Loop
    linhavazia = linhavazia - 1
    ultimalinha = Range("C" & Rows.Count).End(xlUp).Row

    Workbooks(wb).Sheets("Banco de Dados").Range(Cells(linhavazia, 1), Cells(linhavazia, 2)).Select
    Selection.AutoFill Destination:=Range(Cells(linhavazia, 1), Cells(ultimalinha, 2)), Type:=xlFillValues
  
    
End Sub
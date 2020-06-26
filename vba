Sub autoavaliar_1()

    
    nome = Worksheets("Atualizar").Range("C11")
    origem = Sheets("Atualizar").Range("E11") & "\" & Sheets("Atualizar").Range("C7") & "\"
    dia = Sheets("Atualizar").Range("C7")
    
    Filename = origem & nome & " " & dia & ".xls"
    wb = ThisWorkbook.Name
    
    'verifica linhavazia do recipiente para nao sobrepor
    Workbooks(wb).Sheets("Autoavaliar_1").Activate
    linhavazia = Range("A" & Rows.Count).End(xlUp).Row + 1
    
    Workbooks.Open Filename:=Filename
               
''''''''''''''''''
    'tratar planilha
    Rows(1).Select
    Selection.Delete
    
    Cells.Select
    Selection.UnMerge
''''''''''''''''''''
    ultimalinha = Range("A" & Rows.Count).End(xlUp).Row - 1
    
    'Copia as informações do arquivo do NBS e cola na aba de BD Vendedores
    Range("B2:B" & ultimalinha).Select
    Selection.Copy
    Workbooks(wb).Sheets("Autoavaliar_1").Range("A" & linhavazia).PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False
        
    'Copia as informações do arquivo do NBS e cola na aba de BD Vendedores
    Range("D2:K" & ultimalinha).Select
    Selection.Copy
    Workbooks(wb).Sheets("Autoavaliar_1").Range("B" & linhavazia).PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False
    
    'Copia as informações do arquivo do NBS e cola na aba de BD Vendedores
    Range("M2:R" & ultimalinha).Select
    Selection.Copy
    Workbooks(wb).Sheets("Autoavaliar_1").Range("j" & linhavazia).PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False
    
    'Copia as informações do arquivo do NBS e cola na aba de BD Vendedores
    Range("T2:T" & ultimalinha).Select
    Selection.Copy
    Workbooks(wb).Sheets("Autoavaliar_1").Range("p" & linhavazia).PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False
    
 'Salva a planilha gerada do NBS
    Application.DisplayAlerts = False
    Workbooks(nome & " " & dia & ".xls").Close SaveChanges:=False
    Application.DisplayAlerts = True
    
    ActiveWorkbook.Sheets("Autoavaliar_1").Select
    Range("A1").Select
    
    
End Sub

 Sub visualizações()

    
    nome = Worksheets("Atualizar").Range("C12")
    origem = Sheets("Atualizar").Range("E11") & "\" & Sheets("Atualizar").Range("C7") & "\"
    dia = Sheets("Atualizar").Range("C7")
    
    Filename = origem & nome & " " & dia & ".xlsx"
    wb = ThisWorkbook.Name
    
    'verifica linhavazia do recipiente para nao sobrepor
    Workbooks(wb).Sheets("Autoavaliar_2").Activate
    linhavazia = Range("A" & Rows.Count).End(xlUp).Row + 1
        
    Workbooks.Open Filename:=Filename
    ultimalinha = Range("A" & Rows.Count).End(xlUp).Row
    
    'Copia as informações do arquivo do NBS e cola na aba de BD Vendedores
    Range("A2:E" & ultimalinha).Select
    Selection.Copy
    Workbooks(wb).Sheets("Autoavaliar_2").Range("B" & linhavazia).PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False
        
    
 'Salva a planilha gerada do NBS
    Application.DisplayAlerts = False
    Workbooks(nome & " " & dia & ".xlsx").Close SaveChanges:=True
    Application.DisplayAlerts = True
    
    ActiveWorkbook.Sheets("Autoavaliar_2").Select
    Range("A1").Select
    
 'isolar placa
    Columns("A:A").Select
    'Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    'Range("A1").Select
    'ActiveCell.FormulaR1C1 = "Placa"
    Range("A" & linhavazia).Select
    ActiveCell.FormulaR1C1 = "=LEFT(RC[5],7)"
    linhavazia2 = Range("B" & Rows.Count).End(xlUp).Row
    Range("A" & linhavazia).Select
    Selection.AutoFill Destination:=Range(Cells(linhavazia, 1), Cells(linhavazia2, 1))
    
    
End Sub
Sub publicações()

    nome = Worksheets("Atualizar").Range("C13")
    origem = Sheets("Atualizar").Range("E11") & "\" & Sheets("Atualizar").Range("C7") & "\"
    dia = Sheets("Atualizar").Range("C7")
    
    Filename = origem & nome & " " & dia & ".xlsx"
    wb = ThisWorkbook.Name
    
    
    'verifica linhavazia do recipiente para nao sobrepor
    Workbooks(wb).Sheets("Autoavaliar_3").Activate
    linhavazia = Range("A" & Rows.Count).End(xlUp).Row + 1
    
    
    Workbooks.Open Filename:=Filename
    ultimalinha = Range("A" & Rows.Count).End(xlUp).Row
    
    'Copia as informações do arquivo do NBS e cola na aba de BD Vendedores
    Range("C2:H" & ultimalinha).Select
    Selection.Copy
    Workbooks(wb).Sheets("Autoavaliar_3").Range("A" & linhavazia).PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False
        
    
 'Salva a planilha gerada do NBS
    Application.DisplayAlerts = False
    Workbooks(nome & " " & dia & ".xlsx").Close SaveChanges:=True
    Application.DisplayAlerts = True
    
    ActiveWorkbook.Sheets("Autoavaliar_3").Select
    Range("A1").Select
End Sub

Sub nbs_gm()
    nome = Worksheets("Atualizar").Range("C14")
    origem = Sheets("Atualizar").Range("E11") & "\" & Sheets("Atualizar").Range("C7") & "\"
    dia = Sheets("Atualizar").Range("C7")
    
    Filename = origem & nome & " " & dia & ".xlsx"
    wb = ThisWorkbook.Name
         
    Workbooks.Open Filename:=Filename
    
''''''''''''''''''
    'tratar planilha
    Rows(1).Select
    Selection.Delete
    
    Rows(1).Select
    Selection.Delete
''''''''''''''''''''
    ultimalinha = Range("A" & Rows.Count).End(xlUp).Row + 1
    
    'Copia as informações do arquivo do NBS e cola na aba de BD Vendedores
    Range("A2:E" & ultimalinha).Select
    Selection.Copy
        
    Workbooks(wb).Sheets("NBS").Activate
    ultimalinha = Range("A" & Rows.Count).End(xlUp).Row + 1
    Workbooks(wb).Sheets("NBS").Range(Cells(ultimalinha, 1), Cells(ultimalinha, 5)).PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False
        
        
 'Salva a planilha gerada do NBS
    Application.DisplayAlerts = False
    Workbooks(nome & " " & dia & ".xlsx").Close SaveChanges:=False
    Application.DisplayAlerts = True
    
    ActiveWorkbook.Sheets("NBS").Select
    Range("A1").Select
    
    
End Sub

Sub nbs_ni()
    nome = Worksheets("Atualizar").Range("C15")
    origem = Sheets("Atualizar").Range("E11") & "\" & Sheets("Atualizar").Range("C7") & "\"
    dia = Sheets("Atualizar").Range("C7")
    
    Filename = origem & nome & " " & dia & ".xlsx"
    wb = ThisWorkbook.Name
    
    Workbooks.Open Filename:=Filename
     
''''''''''''''''''
    'tratar planilha
    Rows(1).Select
    Selection.Delete
    
    Rows(1).Select
    Selection.Delete
''''''''''''''''''''
    ultimalinha = Range("A" & Rows.Count).End(xlUp).Row + 1
    
    'Copia as informações do arquivo do NBS e cola na aba de BD Vendedores
    Range("A2:E" & ultimalinha).Select
    Selection.Copy
        
    Workbooks(wb).Sheets("NBS").Activate
    ultimalinha = Range("A" & Rows.Count).End(xlUp).Row + 1
    Workbooks(wb).Sheets("NBS").Range(Cells(ultimalinha, 1), Cells(ultimalinha, 5)).PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False
        
        
 'Salva a planilha gerada do NBS
    Application.DisplayAlerts = False
    Workbooks(nome & " " & dia & ".xlsx").Close SaveChanges:=False
    Application.DisplayAlerts = True
    
    ActiveWorkbook.Sheets("NBS").Select
    Range("A1").Select
    
    
End Sub

Sub nbs_vw()
    nome = Worksheets("Atualizar").Range("C16")
    origem = Sheets("Atualizar").Range("E11") & "\" & Sheets("Atualizar").Range("C7") & "\"
    dia = Sheets("Atualizar").Range("C7")
    
    
    Filename = origem & nome & " " & dia & ".xlsx"
    wb = ThisWorkbook.Name
    
    Workbooks.Open Filename:=Filename
    
''''''''''''''''''
    'tratar planilha
    Rows(1).Select
    Selection.Delete
    
    Rows(1).Select
    Selection.Delete
''''''''''''''''''''
    ultimalinha = Range("A" & Rows.Count).End(xlUp).Row + 1
    
    'Copia as informações do arquivo do NBS e cola na aba de BD Vendedores
    Range("A2:E" & ultimalinha).Select
    Selection.Copy
        
    Workbooks(wb).Sheets("NBS").Activate
    ultimalinha = Range("A" & Rows.Count).End(xlUp).Row + 1
    Workbooks(wb).Sheets("NBS").Range(Cells(ultimalinha, 1), Cells(ultimalinha, 5)).PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False
        
        
 'Salva a planilha gerada do NBS
    Application.DisplayAlerts = False
    Workbooks(nome & " " & dia & ".xlsx").Close SaveChanges:=False
    Application.DisplayAlerts = True
    
    ActiveWorkbook.Sheets("NBS").Select
    Range("A1").Select
    
    
End Sub
Sub tratar_placa()
wb = ThisWorkbook.Name
Workbooks(wb).Sheets("NBS").Select

Columns("B:B").Select
    Selection.Replace What:="-", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

End Sub



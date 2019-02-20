Sub OcultarPlanilhas()

Dim xSht As Variant

Dim cont As Long

Dim wkb As Workbook

Dim wks As Worksheet

Set wkb = ThisWorkbook


For Each wks In wkb.Worksheets
    cont = 0
        For Each xSht In ActiveWorkbook.Sheets
            If xSht.Visible Then cont = cont + 1
        Next
 
If cont > 1 Then
    wks.Visible = xlSheetInVisible

Else
    Exit Sub
End If

Next wks

End Sub

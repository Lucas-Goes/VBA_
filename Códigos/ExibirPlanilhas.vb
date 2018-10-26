Sub ExibirPlanilhas()

Dim wkb As Workbook

Dim wks As Worksheet

Dim intUltimoDia As Integer

Set wkb = ThisWorkbook

For Each wks In wkb.Worksheets

wks.Visible = xlSheetVisible

Next wks

End Sub

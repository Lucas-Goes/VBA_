Sub Fazer_Backup()
'Este modo faz o backup da Folha desejada em algum diretório com data e hora

Application.DisplayAlerts = False

Set wsBackup = ThisWorkbook.Sheets("Base_Dados")

wsCadastro.Copy ' Técnica para o backup

destino = "C:\"
data = Left(Date, 2) & "-" & Month(Date) & "-" & Year(Date)
hora = Now()
hora = Mid(hora, 12, 2) & Mid(hora, 15, 2) & Right(hora, 2)

nome = destino & "Lucimara_bckp_Base de Dados Mesa " & data & hora & ".xlsm"

ActiveWorkbook.SaveAs Filename:=nome, _
        FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
ActiveWorkbook.Close

Application.DisplayAlerts = True

End Sub
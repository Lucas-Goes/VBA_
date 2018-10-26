 Sub Carregar_Dados()
'Carrega Dados de outra Planilha utilizando parâmetros em células

    'Busca arquivo pelo nome e caminho informados nos parâmetros   
    'Exemplo
    nome = sheets("Atualizar").Range("D12") 
    origem = Sheets("Atualizar").Range("D10")
    marca = Sheets("Atualizar").Range("D24")
    'Definição do nome 
    filename = origem & nome & " - " & marca & ".xlsx"

    wb = ThisWorkbook.Name
    
    'Abre o arquivo cujo qual queremos as informações
    Workbooks.Open filename:=filename
    'Identifica a quantidade de registros no arquivo
    ultimalinha = Range("A" & Rows.Count).End(xlUp).Row
    'ultimacoluna = Cells(1, Columns.Count).End(xlToLeft).Column

    'Copia as informações do arquivo 
    Range("A2:D" & ultimalinha).Select 'definir última coluna
    Selection.Copy
    
    'Retorna para a planilha que vai receber os dados
    Workbooks(wb).Sheets("Banco de Dados").Activate
    'Verifica qual a última linha vazia para colar as informações 
    linhavazia = 2 'primeira linha a verificar pois normalmente se tem um cabeçalho
    Do While Not IsEmpty(Range("C" & linhavazia))
        linhavazia = linhavazia + 1
    Loop
        Workbooks(wb).Sheets("Banco de Dados").Range(Cells(linhavazia, 3), Cells(linhavazia, 6)).PasteSpecial Paste:=xlPasteValues
    
    'Fecha a planilha com os dados sem salvar
    Application.DisplayAlerts = False
    Workbooks(nome & " - " & marca & ".xlsx").Close SaveChanges:=False
    Application.DisplayAlerts = True
    
  
    End Sub
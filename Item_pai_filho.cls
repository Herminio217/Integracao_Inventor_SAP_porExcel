Sub Item_pai_filho()
' neste caso o w é uma variável para declarar a planilha onde está a informação
Sheets("Plan2").Select
Dim percorrer As Integer
percorrer = 1
Do While Cells(percorrer, 1) <> ""
    If Cells(percorrer, 1) = "Código" Then
        'adiciona 1 a variável para descer uma célula da atual
        percorrer = percorrer + 1
        'copia a célula ativa
        Cells(percorrer, 1).Copy
        'toma como referência a célula ativa e cola na coluna ao lado
        Cells(percorrer, 2).PasteSpecial
        percorrer = percorrer + 1
                
        ' Faça até encontrar o valor Código
        Do Until Cells(percorrer, 1) = "Código" OR Cells(percorrer, 1) = ""
        Cells(percorrer, 1).Copy
        Cells(percorrer, 3).PasteSpecial
        percorrer = percorrer + 1
        ' a única diferença é que cola a célula duas colunas depois, que vai ser a coluna filha
        Loop
    End If
Loop

' fazer código para que ao final deste comando o excel exclua todos as linhas que tem a palavra (código)
' aqui no caso eu vou ter que selecionar a célula A1 primeiro pra depois mandar para baixo e só assim executar o resto do comando 
Range(A1).Offset(-1,0)
ActiveCell.Offset(1, 0).Select
    'ActiveSheet.ListObjects("Teste__2").Range.AutoFilter Field:=1, Criteria1:= _
    '    "Código"
    'Range("A7").Select
    'Range(Selection, Selection.End(xlDown)).Select
    'Selection.EntireRow.Delete
    'ActiveSheet.ListObjects("Teste__2").Range.AutoFilter Field:=1
End Sub    
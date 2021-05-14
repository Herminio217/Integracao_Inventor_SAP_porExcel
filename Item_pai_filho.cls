Sub Item_pai_filho()
' Selecionar planilha Plan2 que é onde será executado o código
Sheets("Plan2").Select

' Selecionar colunas B e C que são a Pai e Filho para converter em texto
Columns("B:C").Select
Selection.NumberFormat = "@"

' Código para colocar o primeiro item na coluna Pai e os demais itens na coluna filho
Dim percorrer As Integer
percorrer = 1
Do While Cells(percorrer, 1) <> ""
    If Cells(percorrer, 1) = "Código" Then
        'adiciona 1 a variável para descer uma célula da atual
        percorrer = percorrer + 1
        'copia a célula ativa
        Cells(percorrer, 2).Value = Cells(percorrer, 1).Value
        'toma como referência a célula ativa e cola na coluna ao lado
        percorrer = percorrer + 1
                
        ' Faça até encontrar o valor Código
        Do Until Cells(percorrer, 1) = "Código" Or Cells(percorrer, 1) = ""
        ' as colunas devem estar configuradas para receber valores do tipo TEXTO
        Cells(percorrer, 3).Value = Cells(percorrer, 1).Value
        percorrer = percorrer + 1
        ' a única diferença é que cola a célula duas colunas depois, que vai ser a coluna filha
        Loop
    End If
Loop

' Código para que ao final deste comando o excel exclua todos as linhas que tem a palavra (código)

' filtrar por "código"
ActiveSheet.ListObjects("Teste__2").Range.AutoFilter Field:=1, Criteria1:="Código"
' selecionar a primeira linha
Rows("1:1").Select
' selecionar uma célula a baixo da referenciada anteriormente
ActiveCell.Offset(1, 0).Select
' selecionar todas as células a baixo até a preenchida
Range(Selection, Selection.End(xlDown)).Select
' deletar a seleção
Selection.EntireRow.Delete
' comando para limpar o filtro
ActiveSheet.ShowAllData

' Igualar Pai Filho. código para preencher a coluna pai de acordo com a coluna filho para depois exportar para a ITT1

    Sheets("Plan2").Select
    'declaração de variável para percorrer as linhas
    Dim lPowerQuery3 As Integer
    ' começa a partir da linha 2 porque na linha 1 está o cabeçalho
    lPowerQuery3 = 2
    ' faça até na coluna pai e na coluna filho não tenham mais valores
    Do Until Cells(lPowerQuery3, 2) = "" And Cells(lPowerQuery3, 3) = ""
        ' se a celular estiver preenchida quer dizer que esse é o código a ser copiado 
        If Cells(lPowerQuery3, 2) <> "" Then
            ' loop para ir copiando o código achado para as linhas em branco até achar o próximo código a ser copiado
            Do While Cells(lPowerQuery3, 2).Offset(1, 0) = "" And Cells(lPowerQuery3,2).Offset(1,1) <> ""
            Cells(lPowerQuery3, 2).Offset(1, 0).Value = Cells(lPowerQuery3, 2).Value
            ' acrimento
            lPowerQuery3 = lPowerQuery3 + 1
            Loop
        End If
        ' acrimento
        lPowerQuery3 = lPowerQuery3 + 1
    Loop
End Sub    
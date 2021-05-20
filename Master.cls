Sub Atualizar_Exportar()
    ' Atualizar os dados da planilhas, trás os dados das planilhas através do Power Query
    ActiveWorkbook.RefreshAll

' ------------------------------------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------------------------

' Item_pai_Filho
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

' ------------------------------------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------------------------

' Dados_para_OITM
    ' Código que traz os dados da coluna CÓDIGO e passa para ITEM CODE
    Sheets("Plan2").Select
    ' Seleciona os dados da coluna A que é onde fica o código
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("OITM - Cadastro do Item").Select
    ' Seleciona a coluna A que é onde fica o itemcode
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
            
    ' Códgio que traz os dados da coluna DESCRIÇÃO e passa para ItemName
    Sheets("Plan2").Select
    ' Seleciona os dados da coluna B que é onde fica a Descrição
    Range("D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("OITM - Cadastro do Item").Select
    ' Seleciona a coluna B que é onde fica o itemname
    Range("B3").Select
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False

    ' Código que prenche as coluna DO TIPO DE ITEM com um valor Y (yes) enquanto outra colina estivar preenchida
    ' Comando para selecionar a tabale aonde será executado o comando
    Sheets("OITM - Cadastro do Item").Select
    ' variável que inicia a linha de verificaçãoi e preenchimento
    linhaY = 3
    ' Estrutura de repetição para enquanto a linha tiver preenchimento ela preencher automaticamente
    While Cells(linhaY, 1) <> ""
        ' se a célula especificada na coluna 1 for diferente de vazio então a célula especificada na coluna 4 será preenchida com Y
        If Cells(linhaY, 1) <> "" Then
            Cells(linhaY, 4) = "Y"
            Cells(linhaY, 5) = "Y"
            Cells(linhaY, 6) = "Y"
            linhaY = linhaY + 1
        End If
    Wend

' ------------------------------------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------------------------

' Dados_para_OITT
    ' selecionar planilha de trabalho
    Sheets("Plan2").Select
    'declaração de variável para percorrer as linhas
    Dim lPowerQuery, lOITT as Integer
    ' começa a partir da linha 2 porque na linha 1 está o cabeçalho
    lPowerQuery = 2
    lOITT = 3
    Do Until Cells(lPowerQuery, 2) = "" And Cells(lPowerQuery,3) = ""
        If Cells(lPowerQuery, 2) <> "" Then
            ' iguala valor do item pai identificado na PowerQuery para a OITT no campo Tree code
            Worksheets("OITT - Estrutura de Produto").Cells(lOITT,1).Value = Worksheets("Plan2").Cells(lPowerQuery, 2).Value
            ' iguala a quantidade do item pai identificado na PowerQuery para a OITT no campo quantity
            Worksheets("OITT - Estrutura de Produto").Cells(lOITT,3).Value = Worksheets("Plan2").Cells(lPowerQuery, 5).Value
            ' toda vez que finalizar acrescenta 1 para que o próximo valor encontrado seja armazenado na linha seguinte na OITT
            lOITT = lOITT + 1
        End if
        ' adiciona 1 para testar a condição de preenchimento na linha seguinte
        lPowerQuery = lPowerQuery + 1
    Loop

    ' Código que prenche as coluna TREETYPE com um valor padrão iProduction enquanto outra colina estivar preenchida
    ' Comando para selecionar a tabale aonde será executado o comando
    Sheets("OITT - Estrutura de Produto").Select
    ' variável que inicia a linha de verificaçãoi e preenchimento
    linha_iP = 3
    ' Estrutura de repetição para enquanto a linha tiver preenchimento ela preencher automaticamente
    While Cells(linha_iP, 1) <> ""
        ' se a célula especificada na coluna 1 for diferente de vazio então a célula especificada na coluna 2 será preenchida com iProduction
        If Cells(linha_iP, 1) <> "" Then
            Cells(linha_iP, 2) = "iProduction"
            linha_iP = linha_iP + 1
        End If
    Wend
    
' ------------------------------------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------------------------

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

' ------------------------------------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------------------------

' Dados_para_ITT1()

    ' converter para texto a linha do código filho e linha número para poder receber os valores corretos
    Worksheets("ITT1 - Campos OITT").Select
    Range("B:B,C:C,E:E").Select
    Selection.NumberFormat = "@"
    ' selecionar planilha de trabalho
    Sheets("Plan2").Select
    ' declaração de variável para percorrer as linhas
    Dim lPowerQuery4, lITT1 As Integer
    ' começa a partir da linha 2 porque na linha 1 está o cabeçalho
    lPowerQuery4 = 2
    lITT1 = 3
    Do Until Cells(lPowerQuery4, 2) = "" And Cells(lPowerQuery4, 3) = ""
        If Cells(lPowerQuery4, 3) <> "" Then
            ' iguala valor Item Pai para ITT1
            Worksheets("ITT1 - Campos OITT").Cells(lITT1, 1).Value = Worksheets("Plan2").Cells(lPowerQuery4, 2).Value
            ' iguala valor Item Filho para ITT1
            Worksheets("ITT1 - Campos OITT").Cells(lITT1, 3).Value = Worksheets("Plan2").Cells(lPowerQuery4, 3).Value
            ' iguala valor Quantidade para ITT1
            Worksheets("ITT1 - Campos OITT").Cells(lITT1, 4).Value = Worksheets("Plan2").Cells(lPowerQuery4, 5).Value
            ' iguala valor Depósito para ITT1
            Worksheets("ITT1 - Campos OITT").Cells(lITT1, 5).Value = Worksheets("Plan2").Cells(lPowerQuery4, 7).Value
        ' toda vez que finalizar acrescenta 1 para que o próximo valor encontrado seja armazenado na linha seguinte na OITT
        lITT1 = lITT1 + 1
        End If
        ' adiciona 1 para testar a condição de preenchimento na linha seguinte
        lPowerQuery4 = lPowerQuery4 + 1
    Loop

    ' código para prenncher a sequência 01, 02, 03... referente as linhas na tabela ITT1
    ' preenche em sequência de acordo com a quantidade de itens
    Sheets("ITT1 - Campos OITT").Select
    linhaLine = 3
    acrimento = 1
    While Cells(linhaLine, 1) <> ""
        ' compara com o valor de baixo, assim que ele
        Do Until Cells(linhaLine, 1) <> Cells(linhaLine, 1).Offset(1, 0)
            Cells(linhaLine, 2) = acrimento
            linhaLine = linhaLine + 1
            acrimento = acrimento + 1
        Loop
        Cells(linhaLine, 2) = acrimento
        linhaLine = linhaLine + 1
        acrimento = 1
    Wend

' ------------------------------------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------------------------
    ' Desativa as notificações após uma ação que no caso seria quando fechar. Quando salvar, caso tenha outro arquivo com mesmo nome essa função
    ' desabilia a text box (deseja substituir arquivo existente?) e permite que salve por cima.
    Application.DisplayAlerts = False
    ' Exportacao
    ' esse código exporta como tabulação a planilha OITM para ser importada pelo SAP via DTW
    ThisWorkbook.Sheets("OITM - Cadastro do Item").Copy
    ' salva o arquivo em uma pasta específica, nesse caso na área de trabalho
    ActiveWorkbook.SaveAs Filename:="C:\Users\Léo Hermínio\Desktop\ OITM DTW.txt", FileFormat:=xlText
    ActiveWorkbook.Close SaveChanges:=True

    ' esse código exporta como tabulação a planilha OITT para ser importada pelo SAP via DTW
    ThisWorkbook.Sheets("OITT - Estrutura de Produto").Copy
    ' salva o arquivo em uma pasta específica, nesse caso na área de trabalho
    ActiveWorkbook.SaveAs Filename:="C:\Users\Léo Hermínio\Desktop\ OITT DTW.txt", FileFormat:=xlText
    ActiveWorkbook.Close SaveChanges:=True

    ' esse código exporta como tabulação a planilha ITT1 para ser importada pelo SAP via DTW
    ThisWorkbook.Sheets("ITT1 - Campos OITT").Copy
    ' salva o arquivo em uma pasta específica, nesse caso na área de trabalho
    ActiveWorkbook.SaveAs Filename:="C:\Users\Léo Hermínio\Desktop\ ITT1 DTW.txt", FileFormat:=xlText
    ActiveWorkbook.Close SaveChanges:=True

    ' Reativa as notificações
    Application.DisplayAlerts = True
' ------------------------------------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------------------------

' código para limpar os dados após a exportação
    Sheets("OITM - Cadastro do Item").Select
    Rows("3:" & Rows.Count).ClearContents

    Sheets("ITT1 - Campos OITT").Select
    Rows("3:" & Rows.Count).ClearContents

    Sheets("OITT - Estrutura de Produto").Select
    Rows("3:" & Rows.Count).ClearContents

' ------------------------------------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------------------------
    ' fechar o form
    Unload Form_Exportacao

End Sub
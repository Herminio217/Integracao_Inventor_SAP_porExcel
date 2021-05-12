Sub Dados_para_OITT
    ' selecionar planilha de trabalho
    Sheets("Plan2").Select
    'declaração de variável para percorrer as linhas
    Dim lPowerQuery, lOITT as Integer
    ' começa a partir da linha 2 porque na linha 1 está o cabeçalho
    lPowerQuery = 2
    lOITT = 4
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
    linha_iP = 4
    ' Estrutura de repetição para enquanto a linha tiver preenchimento ela preencher automaticamente
    While Cells(linha_iP, 1) <> ""
        ' se a célula especificada na coluna 1 for diferente de vazio então a célula especificada na coluna 2 será preenchida com iProduction
        If Cells(linha_iP, 1) <> "" Then
            Cells(linha_iP, 2) = "iProduction"
            linha_iP = linha_iP + 1
        End If
    Wend
End Sub
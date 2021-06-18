Sub Preenchimento_Auto()
' código para prenncher a sequência 01, 02, 03... referente as linhas na tabela ITT1
' preenche em sequência de acordo com a quantidade de itens
Sheets ("ITT1 - Campos OITT").Select
linhaLine = 4
acrimento = 1
While Cells(linhaLine, 5) <> ""
    If Cells(linhaLine, 5) <> "" Then
        Cells(linhaLine, 2) = "0" & acrimento
        linhaLine = linhaLine + 1
        acrimento = acrimento + 1
    End If
Wend

' Código que prenche as coluna DO TIPO DE ITEM com um valor Y (yes) enquanto outra colina estivar preenchida
' Comando para selecionar a tabale aonde será executado o comando
Sheets("OITM - Cadastro do Item").Select
' variável que inicia a linha de verificaçãoi e preenchimento
linhaY = 4
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
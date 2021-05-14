Sub Dados_para_ITT1()
    ' Código para igualar itme pai e item filho
    Sheets("Plan2").Select
    ' declaração de variável para percorrer as linhas
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
    lITT1 = 4
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
    linhaLine = 4
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
End Sub
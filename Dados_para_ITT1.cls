Sub Dados_para_OITM()
    ' converter para texto a linha do código filho e linha número para poder receber os valores corretos
    Worksheets("ITT1 - Campos OITT").Select
    Columns("B:C").Select
    Selection.NumberFormat = "@"
    ' selecionar planilha de trabalho
    Sheets("Plan2").Select
    'declaração de variável para percorrer as linhas
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
        End If
        ' toda vez que finalizar acrescenta 1 para que o próximo valor encontrado seja armazenado na linha seguinte na OITT
        lITT1 = lITT1 + 1
        ' adiciona 1 para testar a condição de preenchimento na linha seguinte
        lPowerQuery4 = lPowerQuery4 + 1
    Loop
End Sub
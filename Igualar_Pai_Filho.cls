Sub Igualar_Pai_Filho()
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
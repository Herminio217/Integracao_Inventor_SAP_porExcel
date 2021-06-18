Sub Dados_para_OITM()
    ' Código que traz os dados da coluna CÓDIGO e passa para ITEM CODE
    Sheets("Plan2").Select
    ' Seleciona os dados da coluna A que é onde fica o código
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("OITM - Cadastro do Item").Select
    ' Seleciona a coluna A que é onde fica o itemcode
    Range("A4").Select
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
    Range("B4").Select
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False

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
End Sub

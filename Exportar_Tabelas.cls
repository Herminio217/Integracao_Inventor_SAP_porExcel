Sub Exportacao()
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
End Sub
   
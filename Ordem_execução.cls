' Power query para trazer todos os dados para apenas uma planilha
' Separar item pai de item filho na planilha do power query e excluir linhas com o nome código (OITT)
' Trazer dados para a (OITM) Item code e Item name (entrar em contato com matheus para ver como ficou a questão do grupo de item)
' Traezr os dados para a (OITT) Tree code que é o item pai, Quantity que é a quantidade e preencher automático iProduction

' Item_pai_filho.cls
' Dados_para_OITM.cls
' Dados_para_OITT.cls
' Dados_para_ITT1.cls

' poderia mandar o EXCEL excluir todas as primerias linhas das planilhas exportadas, já para não haver bronca de ter que tirar, já que
' já está todo feito

' Código para atualizar a planilha
Sub Atualizar()
ActiveWorkbook.RefreshAll
End Sub

' OITM
' ItemCode - nº do item
' ItemName - Descrição do Item
' ItemsGroupCode - Grupo de Itens
' PurchaseItem - Item de Compra
' SalesItem - Item de Venda
' IventoryItem - Item de Estoque
' 
' OITT
' TreeCode - Item Superior (pai)
' TreeType - Estrutura de produto
' Quantity - nº de unidades, quantidade
' 
' ITT1
' ParentKey - Item Superior (pai)
' LineNum - número da linha
' ItemCode - Item Filho
' Quantity - nº de unidades, quantidade
' Warehouse - depósito
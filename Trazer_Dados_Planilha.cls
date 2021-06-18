let
    Fonte = Folder.Files("C:\Users\Léo Hermínio\Documents\Arquivos Pessoais\NAVIGO\NAVIGO\Planilhas Itens A&E\teste a&e\Teste"),
    #"Colunas Renomeadas" = Table.RenameColumns(Fonte,{{"Name", "Pasta de Trabalho"}}),
    #"Linhas Filtradas" = Table.SelectRows(#"Colunas Renomeadas", each [Extension] = ".xlsx"),
    #"Outras Colunas Removidas" = Table.SelectColumns(#"Linhas Filtradas",{"Content", "Pasta de Trabalho"}),
    #"Arquivos Ocultos Filtrados1" = Table.SelectRows(#"Outras Colunas Removidas", each [Attributes]?[Hidden]? <> true),
    #"Invocar Função Personalizada1" = Table.AddColumn(#"Arquivos Ocultos Filtrados1", "Transformar Arquivo de Teste (2)", each #"Transformar Arquivo de Teste (2)"([Content])),
    #"Outras Colunas Removidas1" = Table.SelectColumns(#"Invocar Função Personalizada1", {"Transformar Arquivo de Teste (2)","Pasta de Trabalho"}),
    #"Coluna de Tabela Expandida1" = Table.ExpandTableColumn(#"Outras Colunas Removidas1", "Transformar Arquivo de Teste (2)", Table.ColumnNames(#"Transformar Arquivo de Teste (2)"(#"Arquivo de Amostra (2)"))),
    #"Outras Colunas Removidas2" = Table.SelectColumns(#"Coluna de Tabela Expandida1",{"Name", "Data"}),
    #"Data Expandido" = Table.ExpandTableColumn(#"Outras Colunas Removidas2", "Data", {"Column1", "Column2", "Column3", "Column4", "Column5"}, {"Data.Column1", "Data.Column2", "Data.Column3", "Data.Column4", "Data.Column5"}),
    #"Colunas Removidas" = Table.RemoveColumns(#"Data Expandido",{"Name"}),
    #"Cabeçalhos Promovidos" = Table.PromoteHeaders(#"Colunas Removidas", [PromoteAllScalars=true]),
    #"Tipo Alterado" = Table.TransformColumnTypes(#"Cabeçalhos Promovidos",{{"Código", type text}, {"Descrição", type text}, {"Quantidade", type any}, {"Und", type text}, {"Depósito", type text}}),
    #"Linhas Filtradas1" = Table.SelectRows(#"Tipo Alterado", each [Depósito] <> null),
    #"Coluna Condicional Adicionada" = Table.AddColumn(#"Linhas Filtradas1", "Pai", each if [Código] = "Código" then "" else null),
    #"Coluna Duplicada" = Table.DuplicateColumn(#"Coluna Condicional Adicionada", "Pai", "Pai - Copiar"),
    #"Colunas Renomeadas1" = Table.RenameColumns(#"Coluna Duplicada",{{"Pai - Copiar", "Filho"}}),
    #"Colunas Reordenadas" = Table.ReorderColumns(#"Colunas Renomeadas1",{"Código", "Pai", "Filho", "Descrição", "Quantidade", "Und", "Depósito"})
in
    #"Colunas Reordenadas"
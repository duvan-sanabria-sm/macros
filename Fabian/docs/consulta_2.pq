''
let
    // Convertir columnas necesarias a texto para el join difuso
    Consulta1Text = Table.TransformColumnTypes(Consulta1, {{"query1_netsuite.TRANSACTION_ID", type text}}),
    query3Text = Table.TransformColumnTypes(query3_netsuite, {{"CREATED_FROM_ID", type text}}),

    // Join difuso usando columnas convertidas
    Origen = Table.FuzzyNestedJoin(
        Consulta1Text,
        {"query1_netsuite.TRANSACTION_ID"},
        query3Text,
        {"CREATED_FROM_ID"},
        "Tabla_Consulta_desde_NetSuite24.1",
        JoinKind.FullOuter,
        [IgnoreCase=false, IgnoreSpace=false, NumberOfMatches=1, Threshold=1]
    ),

    // Expandir columnas de la tabla combinada
    #"Se expandió Tabla_Consulta_desde_NetSuite24.1" = Table.ExpandTableColumn(
        Origen,
        "Tabla_Consulta_desde_NetSuite24.1",
        {
            "CREATED_FROM_ID", "TRANID", "STATUS", "AMOUNT_LINKED", "GROSS_AMOUNT",
            "TRANSACTION_TYPE", "CREATE_DATE", "TRANSACTION_ID", "SM_FECHA_REAL_TRANSACCIN"
        },
        {
            "Tabla_Consulta_desde_NetSuite24.1.CREATED_FROM_ID", "Tabla_Consulta_desde_NetSuite24.1.TRANID",
            "Tabla_Consulta_desde_NetSuite24.1.STATUS", "Tabla_Consulta_desde_NetSuite24.1.AMOUNT_LINKED",
            "Tabla_Consulta_desde_NetSuite24.1.GROSS_AMOUNT", "Tabla_Consulta_desde_NetSuite24.1.TRANSACTION_TYPE",
            "Tabla_Consulta_desde_NetSuite24.1.CREATE_DATE", "Tabla_Consulta_desde_NetSuite24.1.TRANSACTION_ID",
            "Tabla_Consulta_desde_NetSuite24.1.SM_FECHA_REAL_TRANSACCIN"
        }
    ),

    // Quitar columna no deseada
    #"Columnas quitadas" = Table.RemoveColumns(#"Se expandió Tabla_Consulta_desde_NetSuite24.1", {"query3_netsuite"})
in
    #"Columnas quitadas"
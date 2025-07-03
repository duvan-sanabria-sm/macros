'
let
    // Combinar query2_netsuite con query1_netsuite usando TRANSACTION_ID y CREATED_FROM_ID
    Join_Query1 = Table.NestedJoin(
        query2_netsuite,
        {"TRANSACTION_ID"},
        query1_netsuite,
        {"CREATED_FROM_ID"},
        "query1_netsuite",
        JoinKind.LeftOuter
    ),

    // Expandir columnas relevantes de query1_netsuite
    Expand_Query1 = Table.ExpandTableColumn(
        Join_Query1,
        "query1_netsuite",
        {"CREATED_FROM_ID", "TRANID", "SM_LOTES", "STATUS", "TRANSACTION_TYPE", "TRANSACTION_ID"},
        {
            "query1_netsuite.CREATED_FROM_ID",
            "query1_netsuite.TRANID",
            "query1_netsuite.SM_LOTES",
            "query1_netsuite.STATUS",
            "query1_netsuite.TRANSACTION_TYPE",
            "query1_netsuite.TRANSACTION_ID"
        }
    ),

    // Combinar con query3_netsuite usando el TRANSACTION_ID expandido de query1_netsuite
    Join_Query3 = Table.NestedJoin(
        Expand_Query1,
        {"query1_netsuite.TRANSACTION_ID"},
        query3_netsuite,
        {"CREATED_FROM_ID"},
        "query3_netsuite",
        JoinKind.LeftOuter
    )

in
    Join_Query3

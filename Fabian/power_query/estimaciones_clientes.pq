let
    Origen = Odbc.Query("dsn=Netsuite2", "SELECT
    t.tranId AS TRANID,
    t.title AS TITLE,
    e.fullName AS FULL_NAME,
    s.entityid AS FULL_NAME2,
    l.netAmount AS AMOUNT,
    CASE 
        WHEN t.type = 'Estimate' THEN 'Estimación'
    ELSE t.type
    END AS TRANSACTION_TYPE,
    t.tranDate AS CREATE_DATE,
    t.id AS TRANSACTION_ID
    
    FROM transaction AS t
    INNER JOIN transactionLine AS l ON t.id = l.transaction
    LEFT JOIN entity AS e ON t.entity = e.id
    LEFT JOIN employee AS s ON t.employee = s.id
    
    WHERE 
            t.type = 'Estimate'
            AND l.netAmount > 0
            AND t.tranDate >= {ts '2025-01-01 00:00:00'}
            AND t.tranDate < {ts '2026-01-01 00:00:00'}
            
    GROUP BY
        t.tranId,
        t.title,
        e.fullName,
        s.entityid,
        l.netAmount,
        t.type,
        t.tranDate,
        t.id
    ")
in
    Origen
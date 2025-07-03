let
    Origen = Odbc.Query("dsn=Netsuite2", "SELECT
    l.createdFrom AS CREATED_FROM_ID,
    t.tranId AS TRANID,
        CASE
            WHEN t.status = 'A' THEN 'Abierta'
            WHEN t.status = 'B' THEN 'Pagado por completo'
            ELSE t.status
            END AS STATUS,
    tal.amountPaid AS AMOUNT_LINKED,
    l.netAmount AS GROSS_AMOUNT,
    CASE 
        WHEN t.type = 'CustInvc' THEN 'Factura de venta'
    ELSE t.type
    END AS TRANSACTION_TYPE,
    t.tranDate AS CREATE_DATE,
    t.id AS TRANSACTION_ID

    FROM transaction AS t
    
    INNER JOIN transactionLine AS l ON t.id = l.transaction
    INNER JOIN TransactionAccountingLine AS tal ON tal.transaction = l.transaction AND tal.transactionline = l.id

    WHERE
        t.type = 'CustInvc'
        AND l.netAmount > 0
        AND t.tranDate >= {ts '2025-01-01 00:00:00'}
        AND t.tranDate < {ts '2026-01-01 00:00:00'}

    GROUP BY
        l.createdFrom,
        t.tranId,
        t.status,
        tal.amountPaid,
        l.netAmount,
        t.type,
        t.tranDate,
        t.id
    ")
in
    Origen
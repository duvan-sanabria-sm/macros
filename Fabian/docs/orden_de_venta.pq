let
    Origen = Odbc.Query("dsn=Netsuite2", "SELECT
    l.createdFrom AS CREATED_FROM_ID,
    t.tranId AS TRANID,
    t.custbody_ks_sm_batch AS SM_LOTES,
       CASE
        WHEN t.status = 'A' THEN 'Aprobación pendiente'
        WHEN t.status = 'B' THEN 'Ejecución de la orden pendiente'
        WHEN t.status = 'C' THEN 'Cancelada'
        WHEN t.status = 'D' THEN 'Parcialmente ejecutada'
        WHEN t.status = 'E' THEN 'Facturación pendiente/parcialmente ejecutada'
        WHEN t.status = 'F' THEN 'Facturación pendiente'
        WHEN t.status = 'G' THEN 'Facturada'
        WHEN t.status = 'H' THEN 'Cerrada'
    ELSE t.status
    END AS STATUS,
    CASE 
        WHEN t.type = 'SalesOrd' THEN 'Orden de venta'
    ELSE t.type
    END AS TRANSACTION_TYPE,
    t.id AS TRANSACTION_ID

    FROM transaction AS t
    INNER JOIN transactionLine AS l ON t.id = l.transaction

    WHERE
        t.type = 'SalesOrd'
        AND t.tranDate >= {ts '2025-01-01 00:00:00'}
        AND t.tranDate < {ts '2026-01-01 00:00:00'}

    GROUP BY
    l.createdFrom,
    t.tranId,
    t.custbody_ks_sm_batch,
    t.status,
    t.type,
    t.id
    
    ")
in
    Origen
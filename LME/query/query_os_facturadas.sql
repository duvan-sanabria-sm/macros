SELECT  
    t.id AS TRANSACTION_ID,
    t.tranId AS TRANID,
    CASE 
        WHEN t.type = 'CustInvc' THEN 'Factura de venta'
        WHEN t.type = 'SalesOrd' THEN 'Orden de venta'
        ELSE t.type
    END AS TRANSACTION_TYPE,
    t.tranDate AS CREATE_DATE
FROM transaction t
WHERE 
    t.type = 'SalesOrd'
    AND t.id IN (
        SELECT l.createdfrom
        FROM transactionLine l
        INNER JOIN transaction t2 ON l.transaction = t2.id
        WHERE t2.type = 'CustInvc'
              AND l.createdfrom IS NOT NULL
    )
GROUP BY 
    t.id, 
    t.tranId,
    t.type,
    t.tranDate

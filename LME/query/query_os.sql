SELECT 
            t.id AS TRANSACTION_ID,
            
              CASE 
                WHEN t.type = 'CustInvc' THEN 'Factura de venta'
                WHEN t.type = 'SalesOrd' THEN 'Orden de venta'
                ELSE t.type
            END AS TRANSACTION_TYPE,
           
           l.createdfrom AS CREATED_FROM_ID,
           t.tranDate AS CREATE_DATE,
           s.entityid AS FULL_NAME,
           t.tranId AS TRANID,

            CASE
                WHEN t.status = 'A' THEN 'Abierta'
                WHEN t.status = 'B' THEN 'Pagado por completo'
                ELSE t.status
            END AS ESTADO

        FROM transaction AS t
        INNER JOIN transactionLine AS l ON t.id = l.transaction
        LEFT JOIN department AS d ON l.department = d.id
        LEFT JOIN employee AS s ON t.employee = s.id

        WHERE 
            t.type = 'CustInvc'
            AND l.createdfrom IN (
                SELECT t.id
                FROM transaction t
                WHERE t.type = 'SalesOrd'
            )

        GROUP BY 
            t.id,
            t.tranId,
            l.createdfrom,
            t.tranDate,
            s.entityid,
            t.type,
            t.status
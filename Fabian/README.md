# 🧾 Descripción del uso de la Macro

La macro `BuscarFacturasLMA` busca y cruza información entre un archivo de datos llamado **"Datos Netsuite 2.xlsx"** y otro archivo activo que contiene varias hojas con registros de **órdenes de servicio y facturas**.  
Su objetivo principal es **llenar automáticamente datos faltantes** (como número de factura, estado, comercial y orden relacionada), y generar un **reporte en un archivo aparte** con los registros que no se pudieron completar correctamente.

---

# ✅ ¿Por qué es útil esta macro?

Esta macro es muy útil para:

- 🔎 Auditoría o seguimiento de órdenes y facturas.  
- ⚙️ Automatización de procesos administrativos que implican cruce de datos.  
- 🚨 Identificación rápida de registros incompletos o erróneos.  
- 📤 Generación de reportes para revisión o envío a terceros (comerciales, financieros, etc).

# 📊 Consultas SQL utilizadas

Las consultas SQL se encuentran en la hoja `datos_netsuite2` del archivo **Datos Netsuite 2.xlsx**, que contiene información exportada desde NetSuite.  
A continuación se listan las principales consultas usadas:

<details>
<summary>🗂 Consulta: Prog en Sitio (Query1)</summary>

```sql
'Consulta basada en Netsuite2.com'
SELECT
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

'Consulta basada en Netsuite.com'
SELECT 
    TRANSACTIONS.CREATED_FROM_ID, 
    TRANSACTIONS.TRANID, 
    TRANSACTIONS.SM_LOTES,
    TRANSACTIONS.STATUS,
    TRANSACTIONS.TRANSACTION_TYPE, 
    TRANSACTIONS.TRANSACTION_ID

FROM "SERVIMETERS S_A_S".Administrador.TRANSACTIONS TRANSACTIONS

WHERE (TRANSACTIONS.TRANSACTION_TYPE='Orden de venta')

GROUP BY 
    TRANSACTIONS.CREATED_FROM_ID, 
    TRANSACTIONS.TRANID, 
    TRANSACTIONS.SM_LOTES, 
    TRANSACTIONS.STATUS,
    TRANSACTIONS.TRANSACTION_TYPE, 
    TRANSACTIONS.TRANSACTION_ID
```
</details>
<details>
<summary>🗂 Consulta: Prog en Sitio (Query2)</summary>

```sql
'Consulta basada en Netsuite2.com'
SELECT
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

'Consulta basada en Netsuite.com'
SELECT 
    TRANSACTIONS.TRANID, 
    TRANSACTIONS.TITLE, 
    ENTITY.FULL_NAME, 
    EMPLOYEES.FULL_NAME, 
    TRANSACTION_LINES.AMOUNT, 
    TRANSACTIONS.TRANSACTION_TYPE,
    TRANSACTIONS.CREATE_DATE, 
    TRANSACTIONS.TRANSACTION_ID

    FROM "SERVIMETERS S_A_S".Administrador.EMPLOYEES EMPLOYEES, "SERVIMETERS S_A_S".Administrador.ENTITY ENTITY, "SERVIMETERS S_A_S".Administrador.TRANSACTION_LINES TRANSACTION_LINES, "SERVIMETERS S_A_S".Administrador.TRANSACTIONS TRANSACTIONS

    WHERE EMPLOYEES.EMPLOYEE_ID = TRANSACTIONS.SALES_REP_ID 
        AND ENTITY.ENTITY_ID = TRANSACTIONS.ENTITY_ID 
        AND TRANSACTIONS.TRANSACTION_ID = TRANSACTION_LINES.TRANSACTION_ID 
        AND ((TRANSACTIONS.TRANSACTION_TYPE='Estimación') 
        AND (TRANSACTION_LINES.AMOUNT>0))
    
    GROUP BY 
        TRANSACTIONS.TRANID, 
        TRANSACTIONS.TITLE, 
        ENTITY.FULL_NAME, 
        EMPLOYEES.FULL_NAME, 
        TRANSACTION_LINES.AMOUNT, 
        TRANSACTIONS.TRANSACTION_TYPE,
        TRANSACTIONS.CREATE_DATE, 
        TRANSACTIONS.TRANSACTION_ID


```
</details>
<details>
<summary>🗂 Consulta: Prog en Sitio (Query3)</summary>

```sql
'Consulta basada en Netsuite2.com'
SELECT
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

'Consulta basada en Netsuite.com'
SELECT 
    TRANSACTIONS.CREATED_FROM_ID, 
    TRANSACTIONS.TRANID, 
    TRANSACTIONS.STATUS, 
    TRANSACTION_LINES.AMOUNT_LINKED, 
    TRANSACTION_LINES.GROSS_AMOUNT, 
    TRANSACTIONS.TRANSACTION_TYPE, 
    TRANSACTIONS.CREATE_DATE, 
    TRANSACTIONS.TRANSACTION_ID, 
    TRANSACTIONS.SM_FECHA_REAL_TRANSACCIN

    FROM "SERVIMETERS S_A_S".Administrador.TRANSACTION_LINES TRANSACTION_LINES, 
    "SERVIMETERS S_A_S".Administrador.TRANSACTIONS TRANSACTIONS

    WHERE 
    TRANSACTIONS.TRANSACTION_ID = TRANSACTION_LINES.TRANSACTION_ID 
    AND ((TRANSACTIONS.TRANSACTION_TYPE='Factura de venta'))

    GROUP BY 
        TRANSACTIONS.CREATED_FROM_ID, 
        TRANSACTIONS.TRANID, 
        TRANSACTIONS.STATUS, 
        TRANSACTION_LINES.AMOUNT_LINKED,
        TRANSACTION_LINES.GROSS_AMOUNT, 
        TRANSACTIONS.TRANSACTION_TYPE, 
        TRANSACTIONS.CREATE_DATE, 
        TRANSACTIONS.TRANSACTION_ID, 
        TRANSACTIONS.SM_FECHA_REAL_TRANSACCIN

    HAVING (TRANSACTION_LINES.GROSS_AMOUNT>0)
```
</details>

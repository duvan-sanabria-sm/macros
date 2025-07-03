# 🎯 ¿Qué hace esta macro?

La macro extrae información de **cotizaciones, pedidos y facturas** desde una hoja de Excel (`Prog en Sitio.xlsx`, conectada a NetSuite), cruzándola con datos registrados manualmente en otra hoja (`PROGRAMACION EN SITIO 2022.xlsx`), y genera automáticamente un **reporte estructurado** en un nuevo archivo Excel.

# 💡 ¿Por qué es útil?

Porque:

- ⚙️ Automatiza un proceso tedioso y repetitivo de comparación y consolidación de datos.
- 🔄 Permite **cruzar múltiples valores por celda** (por ejemplo, varios pedidos o facturas separados por `/`, `,` o `|`).
- 📊 Genera un archivo final con toda la **información clave de seguimiento** de la programación en sitio:
  - Cotizaciones  
  - Valores cotizados y facturados  
  - Estado del pedido  
  - Estado de la factura  

# 📊 Consultas SQL utilizadas

Las consultas SQL se encuentran en la hoja `datos_netsuite2` del archivo **Datos Netsuite 2.xlsx**, que contiene información exportada desde NetSuite.  
A continuación se listan las principales consultas usadas:

<details>
<summary>🗂 Consulta: Ordenes de Venta</summary>

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
<summary>🗂 Consulta: Estimaciones Clientes</summary>

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
<summary>🗂 Consulta: Ventas Facturadas</summary>

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

---

# ⚙️ Consultas Power Query

Las siguientes consultas se encuentran en el directorio:

> 📁 `Fabian/power_query/`

Estas transformaciones extraen y combinan datos desde NetSuite para análisis y generación de reportes.

---

### 🧾 `ventas_facturadas.pq`

📌 **Descripción:**  
Consulta que obtiene todas las **facturas de venta** registradas en NetSuite, con información detallada sobre valor facturado, estado y fecha de emisión.

---

### 📄 `ordenes_de_venta.pq`

📌 **Descripción:**  
Consulta que trae todas las **órdenes de venta** generadas, vinculadas a lotes, clientes, estado y tipo de transacción.

---

### 📐 `estimaciones_clientes.pq`

📌 **Descripción:**  
Filtra y transforma todas las **estimaciones activas** asociadas a clientes, enfocándose en aquellas con valor positivo y dentro del rango del año actual.

---

### 🧩 `cruce_facturas_por_ordenes.pq`

📌 **Descripción:**  
Consulta final que **combina** los resultados de `ventas_facturadas.pq`, `ordenes_de_venta.pq` y `estimaciones_clientes.pq` mediante relaciones entre lotes y órdenes para construir el **reporte consolidado**.

---




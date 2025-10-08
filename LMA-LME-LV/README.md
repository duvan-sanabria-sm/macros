
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

---

# ⚠️ Cambios que debes tener en cuenta en el código

> Estos valores se deben modificar según el tipo de macro (LMA, LME, LV) o según tu entorno de trabajo:

### 🔹 Nombre de la función según el tipo de macro:
```vba
Sub BuscarFacturasLMA()  ' Cambiar por BuscarFacturasLME o BuscarFacturasLV según el caso
```

### 🔹 Ruta donde se genera el archivo de resultados:
```vba
With libroReporte.SaveAs Filename:= _
"C:\Users\duvan.sanabria\OneDrive - Servimeters\Documentos\Macros\LME\resultados\LMA Reporte Facturas " & _
Format(Now(), "DD-MMM-YYYY hh mm AMPM") & ".xlsx"
```

### 🔹 Archivo que contiene las órdenes exportadas desde NetSuite:
```vba
Set excelFacturas = Workbooks("Datos Netsuite 2.xlsx")
' O abrir directamente si no está abierto:
Set excelFacturas = Workbooks.Open("C:\Users\duvan.sanabria\OneDrive - Servimeters\Documentos\Macros\LME\Datos Netsuite 2.xlsx")
```

### 🔹 Archivo activo que se desea actualizar:
```vba
Set excelActualizar = Workbooks("LMA 2025 GMM-RG-54 CONT SEGUI EQUIP V1.xlsx")
```

---

# 📊 Consultas SQL utilizadas

Las consultas SQL se encuentran en la hoja `datos_netsuite2` del archivo **Datos Netsuite 2.xlsx**, que contiene información exportada desde NetSuite.  
A continuación se listan las principales consultas usadas:

<details>
<summary>🗂 Consulta: Hoja F DE OS (Netsuite 2)</summary>

```sql
-- Consulta basada en Netsuite2.com
SELECT 
    t.id AS ID_Transaccion,
    CASE 
        WHEN t.type = 'CustInvc' THEN 'Factura de venta'
        WHEN t.type = 'SalesOrd' THEN 'Orden de venta'
        ELSE t.type
    END AS Tipo_Transaccion,
    
    l.createdfrom AS ID_Orden_Relacionada,
    t.tranDate AS Fecha_Creacion,
    s.entityid AS Nombre_Comercial,
    t.tranId AS Codigo_Transaccion,
    
    CASE
        WHEN t.status = 'A' THEN 'Abierta'
        WHEN t.status = 'B' THEN 'Pagado por completo'
        ELSE t.status
    END AS Estado_Factura

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
        AND t.tranDate >= {ts '2025-01-01 00:00:00'}
        AND t.tranDate < {ts '2026-01-01 00:00:00'}
    )

GROUP BY 
    t.id, t.tranId, l.createdfrom, t.tranDate, s.entityid, t.type, t.status
```

</details>

<details>
<summary>🗂 Consulta Hoja F DE OS (Netsuite 1)</summary>

```sql
-- Consulta basada en Netsuite.com
SELECT 
    TRANSACTIONS.TRANSACTION_ID, 
    TRANSACTIONS.TRANID, 
    TRANSACTIONS.TRANSACTION_TYPE, 
    TRANSACTIONS.CREATED_FROM_ID, 
    TRANSACTIONS.CREATE_DATE, 
    TRANSACTIONS.STATUS, 
    EMPLOYEES.FULL_NAME
FROM 
    "SERVIMETERS S_A_S".Administrador.EMPLOYEES EMPLOYEES,
    "SERVIMETERS S_A_S".Administrador.TRANSACTIONS TRANSACTIONS
WHERE 
    EMPLOYEES.EMPLOYEE_ID = TRANSACTIONS.SALES_REP_ID
    AND TRANSACTIONS.TRANSACTION_TYPE = 'Factura de venta'
    AND TRANSACTIONS.CREATED_FROM_ID IN (
        SELECT TRANSACTIONS.TRANSACTION_ID
        FROM "SERVIMETERS S_A_S".Administrador.TRANSACTIONS TRANSACTIONS
        WHERE TRANSACTIONS.TRANSACTION_TYPE = 'Orden de venta'
    )
GROUP BY 
    TRANSACTIONS.TRANSACTION_ID, 
    TRANSACTIONS.TRANID, 
    TRANSACTIONS.TRANSACTION_TYPE, 
    TRANSACTIONS.CREATED_FROM_ID, 
    TRANSACTIONS.CREATE_DATE, 
    TRANSACTIONS.STATUS, 
    EMPLOYEES.FULL_NAME
```

</details>
<details>
<summary>🗂 Consulta Hoja OS FACTURADAS (Netsuite 2)</summary>

```sql
SELECT  
    t.id AS ID_Orden_Venta,
    t.tranId AS Codigo_Orden_Venta,
    CASE 
        WHEN t.type = 'CustInvc' THEN 'Factura de venta'
        WHEN t.type = 'SalesOrd' THEN 'Orden de venta'
        ELSE t.type
    END AS Tipo_Transaccion,
    t.tranDate AS Fecha_Creacion_Orden
FROM transaction t
WHERE 
    t.type = 'SalesOrd'
    AND t.id IN (
        SELECT l.createdfrom
        FROM transactionLine l
        INNER JOIN transaction t2 ON l.transaction = t2.id
        WHERE t2.type = 'CustInvc'
              AND l.createdfrom IS NOT NULL
              AND t2.tranDate >= {ts '2025-01-01 00:00:00'}
              AND t2.tranDate < {ts '2026-01-01 00:00:00'}
    )
GROUP BY 
    t.id, 
    t.tranId,
    t.type,
    t.tranDate
```

</details>
</details>
<details>
<summary>🗂 Consulta Hoja OS FACTURADAS (Netsuite 1)</summary>

```sql
SELECT 
    TRANSACTIONS.TRANSACTION_ID, 
    TRANSACTIONS.TRANID, 
    TRANSACTIONS.TRANSACTION_TYPE, 
    TRANSACTIONS.CREATE_DATE
    
    FROM 
        "SERVIMETERS S_A_S".Administrador.TRANSACTIONS TRANSACTIONS 
        
    WHERE 
        (TRANSACTIONS.TRANSACTION_TYPE='Orden de venta') 
        AND TRANSACTIONS.TRANSACTION_ID 
        IN (SELECT TRANSACTIONS.CREATED_FROM_ID 
         
    FROM "SERVIMETERS S_A_S".Administrador.TRANSACTIONS TRANSACTIONS 
         
    WHERE (TRANSACTIONS.TRANSACTION_TYPE='Factura de venta')) 
         
    GROUP BY 
        TRANSACTIONS.TRANSACTION_ID, 
        TRANSACTIONS.TRANID, 
        TRANSACTIONS.TRANSACTION_TYPE, 
        TRANSACTIONS.CREATE_DATE
```
</details>

## 📁 Otras hojas consultadas

- **FV DE OS**: también extraída de NetSuite, ubicada en la misma carpeta del archivo de datos.
- **datos_netsuite2**: hoja general que concentra las transacciones.

---

# 📌 Recomendaciones finales

- ✅ Verifica que el archivo `Datos Netsuite 2.xlsx` esté actualizado antes de correr la macro.
- ✅ Cambia el nombre de la macro según el tipo de informe (LMA/LME/LV).
- ✅ Guarda el archivo original antes de ejecutar la macro para evitar sobrescribir datos por error.
- ✅ Documenta cualquier cambio en los nombres de hojas o columnas del archivo fuente para evitar errores en futuras ejecuciones.

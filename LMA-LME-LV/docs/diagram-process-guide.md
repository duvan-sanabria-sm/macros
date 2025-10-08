# 🧾 Explicación Lógica del Proceso: BuscarFacturasLV (Macro VBA)

Este documento describe **paso a paso y con claridad** cómo funciona la macro `BuscarFacturasLV`, qué archivos accede y cómo decide qué hacer con cada dato.

---

## ✅ Flujo General del Proceso

### 🔄 Inicio

1. **Desactiva la actualización de pantalla** para mejorar el rendimiento.
2. Verifica si el archivo `Datos Netsuite 2.xlsx` está abierto:
   - ✅ **Sí está abierto**: simplemente lo usa.
   - ❌ **No está abierto**: lo abre desde una ruta local.
3. Crea un archivo nuevo de resultados llamado algo como:  
   `LMA Reporte Facturas 03-Jul-2025.xlsx`.
4. Abre el archivo principal `LMA 2025 GMM-RG-54 CONT SEGUI EQUIP V1.xlsx`, que contiene varias hojas (una por cada mes).

---

## 🔁 Por Cada Hoja del Archivo LMA

1. Entra a la hoja del mes y recorre las filas desde la celda `N15` hacia abajo.
2. Para cada fila, **lee los siguientes campos**:
   - **Columna N**: Orden de servicio (OS).
   - **Columna O**: Factura de venta.

---

## ⚠️ Condición: OS inválida ("PEDIDO" o "N/A")

- Si la OS está marcada como "PEDIDO" o "N/A", **no se procesa** esa fila.

---

## 🔍 Búsqueda por OS (Orden de Servicio)

1. Busca la OS en la hoja **`OS FACTURADA`** del archivo `Datos Netsuite 2.xlsx`.
2. Si la encuentra:
   - Obtiene el número de **factura de venta** asociado.
   - Luego busca esa factura en la hoja **`FV DE OS`**.
   - Si la factura existe:
     - Extrae el **estado** y el **comercial**.
     - Actualiza el archivo LMA:
       - Columna O: Factura
       - Columna P: Estado
       - Columna Q: Comercial

---

## 🔍 Búsqueda por Factura

1. Si no encuentra la OS, busca directamente la factura en **`FV DE OS`**.
2. Si encuentra la factura:
   - Extrae la OS, estado y comercial.
   - Actualiza en el archivo LMA:
     - Columna N: OS
     - Columna P: Estado
     - Columna Q: Comercial

---

## ❌ Si no encuentra nada

- La fila se **marca como error**.
- Copia las columnas: A, B, G, N, M, O.
- Registra esta fila en el archivo de resultados y anota el **nombre de la hoja**.

---

## 📋 Finalizando

- Después de procesar cada hoja:
  - Copia las filas con error al libro de resultados.
- Al terminar todas las hojas:
  - Cierra el archivo de Netsuite.
  - Guarda y cierra el archivo de resultados.
  - Vuelve a activar la pantalla.

---

## 🧠 Resumen de la lógica con tabla

| Situación         | ¿Qué busca primero? | ¿Dónde lo busca?           | ¿Qué actualiza?      |
|------------------|----------------------|-----------------------------|-----------------------|
| Tengo una OS     | Factura              | `OS FACTURADA` → `FV DE OS`| O (factura), P, Q     |
| Tengo solo factura | OS                | `FV DE OS`                  | N (OS), P, Q          |
| Nada encontrado   | —                   | —                           | Se guarda en errores  |

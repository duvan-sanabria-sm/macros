# ⚙️ Proyecto Dictámenes Servimeters

El proyecto **Dictámenes** está diseñado para **automatizar y estandarizar el proceso de generación, firma y registro de dictámenes eléctricos** (RETIE / RETILAP / RITEL) en Servimeters.  
Este sistema permite gestionar desde la creación del archivo base hasta la carga de la información hacia la plataforma **SICERCO**, garantizando consistencia y trazabilidad en todos los documentos emitidos.

---

## 🧾 Descripción general

Actualmente, Servimeters se encuentra acreditado desde **2014** y debe mantener control sobre los **dictámenes técnicos** generados para sus clientes.  
Para ello, el proceso involucra macros en Excel, hojas de control y formatos normalizados que aseguran la correcta numeración, firma, conversión a PDF y registro en los formularios requeridos por el ONAC y SICERCO.

---

## 🔁 Flujo general del proceso

1. **Recepción de carpetas desde los ingenieros:**  
   Cada carpeta contiene una hoja de control (`Control de inspección.xlsx`) y los archivos de dictamen (`Dictamen 1.xls`, `Dictamen 2.xls`, etc.).

2. **Ingreso de datos y verificación:**  
   La persona encargada abre los archivos y ejecuta una macro para insertar **fecha, lugar y número de dictamen**.  
   Si hay varios archivos, se usa la macro **"Consecutivo"** para numerarlos automáticamente.

3. **Aplicación de firma:**  
   Se ejecuta la macro **"Firma"** que replica la firma digital en todas las hojas correspondientes del dictamen.

4. **Generación de PDF:**  
   El archivo final se exporta con formato institucional, logos y datos ajustados.

5. **Extracción y consolidación de datos:**  
   A través del archivo **"Extractor"**, se toma la información de la hoja de control y se transfiere al registro de dictámenes.

6. **Generación de formato SICERCO:**  
   El sistema genera un archivo estandarizado con los campos requeridos para la carga en la plataforma de **SICERCO**.

---

## 📚 Documentación rápida

- [📋 01. Flujo general del proceso](docs/flujo_plantuml.puml)
- [🔧 02. Manual de uso del sistema](docs/manual_usuario.md)
- [📈 03. Requisitos funcionales](docs/requisitos_funcionales.md)
- [📉 04. Requisitos no funcionales](docs/requisitos_no_funcionales.md)
- [🧮 05. Estructura de macros VBA](src/macros/)
- [💾 06. Ejemplos de salidas (PDF / SICERCO)](outputs/)
- [✅ Checklist de publicación y control](docs/checklist_publicacion.md)

---

## 🧩 Tecnologías y herramientas usadas

- 💻 **Microsoft Excel + VBA (Macros)**  
  - `PERSONAL.XLSB` → Contiene macros principales  
  - Macros: `GenerarDictamen`, `Consecutivo`, `Firma`, `Extractor`

- 🧰 **PlantUML**  
  - Diagrama del flujo completo (`docs/flujo_plantuml.puml`)

- 📂 **GitHub / Git**  
  - Control de versiones y respaldo de macros, plantillas y documentación

- 📑 **Formatos Excel**  
  - `dictamen_base.xlsx`  
  - `extractor.xlsx`  
  - `hoja_control.xlsx`  
  - `sicerco_formato.xlsx`

---

## 🗂️ Estructura del repositorio

```plaintext
dictamenes/
│
├── 📘 README.md
│
├── 📂 src/
│   ├── macros/
│   │   ├── generar_dictamen.bas
│   │   ├── consecutivo.bas
│   │   ├── firma.bas
│   │   └── extractor.bas
│   └── utils/
│       └── helpers_vba.bas
│
├── 📂 docs/
│   ├── flujo_plantuml.puml
│   ├── manual_usuario.md
│   ├── requisitos_funcionales.md
│   ├── requisitos_no_funcionales.md
│   └── checklist_publicacion.md
│
├── 📂 formatos/
│   ├── dictamen_base.xlsx
│   ├── extractor.xlsx
│   ├── hoja_control.xlsx
│   └── sincerco_formato.xlsx
│
└── 📂 outputs/
    ├── ejemplo_dictamen.pdf
    ├── ejemplo_registro.xlsx
    └── ejemplo_sicerco.xlsx

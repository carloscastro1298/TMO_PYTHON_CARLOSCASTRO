# 📊 BPO Decision Making Automation

Sistema automatizado en Python diseñado para analizar, auditar, clasificar y reportar decisiones operativas en entornos **BPO (Business Process Outsourcing)**. Este proyecto reemplaza flujos manuales o basados en herramientas de nodos (como n8n) mediante el procesamiento de archivos Excel, evaluación de métricas de calidad y **TMO (Tiempo Medio de Operación)**, generación de análisis con inteligencia artificial mediante **Google Gemini** y distribución automatizada de reportes por correo electrónico (en formato HTML y archivos Excel adjuntos).

---

## 📋 Tabla de Contenidos

- [Características Principales](#-características-principales)
- [Arquitectura y Flujo de Trabajo](#-arquitectura-y-flujo-de-trabajo)
- [Requisitos Previos e Instalación](#-requisitos-previos-e-instalación)
- [Configuración de Variables de Entorno](#-configuración-de-variables-de-entorno)
- [Estructura del Archivo Excel de Entrada](#-estructura-del-archivo-excel-de-entrada)
- [Guía de Uso](#-guía-de-uso)
- [Reglas de Negocio, Calidad y TMO](#-reglas-de-negocio-calidad-y-tmo)
- [Estructura del Proyecto](#-estructura-del-proyecto)
- [Manejo de Errores y Solución de Problemas](#-manejo-de-errores-y-solución-de-problemas)

---

## 🚀 Características Principales

* **Motor de Clasificación Basado en Reglas:** Evalúa cada registro según la decisión operativa tomada (`Decision`), asignando automáticamente el tipo de gestión (`Decision_Type`), el nivel de riesgo (`Risk_Level`) y la criticidad (`Critical` vs. `Non-Critical`).
* **Auditoría de Efectividad (`Decision_Quality`):** Cruza el nivel de riesgo con el resultado obtenido (`Result`) para determinar si la gestión fue **Efectiva (`Effective`)** o si **Requiere Revisión (`Needs Review`)**.
* **Control de TMO y Sistema de Alertas:** Analiza el tiempo de resolución en minutos (`Resolution_Time_Minutes`), clasificando el rendimiento operativo y activando una bandera de alerta (`Alert_Required = True`) cuando el caso supera los 30 minutos.
* **Análisis Inteligente con Google Gemini:** Integra el SDK oficial `google-genai` (modelo `gemini-2.5-flash`) para redactar diagnósticos concisos en inglés básico con recomendaciones accionables adaptadas a la criticidad de cada caso.
* **Exportación Segmentada en Excel:** Genera automáticamente tres libros de Excel dentro del directorio de salida:
  * `Critical_Cases_Report.xlsx`: Contiene exclusivamente los casos clasificados como críticos (`Risk_Level == "High"`).
  * `Non_Critical_Cases_Report.xlsx`: Contiene los casos de riesgo bajo y medio.
  * `BPO_Complete_Analysis.xlsx`: Consolidado general con todos los casos procesados, métricas calculadas y recomendaciones de la IA.
* **Doble Modalidad de Envío por Correo (SMTP):**
  1. **Resúmenes Visuales en HTML (`bpo_decision_making.py`):** Envía correos con diseño HTML responsivo y tarjetas diferenciadas por color (rojo para críticos, verde para no críticos) con el detalle y análisis de cada caso.
  2. **Envío Masivo de Adjuntos Excel (`config_email.py`):** Script dedicado que valida la existencia de los tres reportes generados y los envía como archivos `.xlsx` adjuntos en correos independientes.

---

## 🔄 Arquitectura y Flujo de Trabajo

```text
[Archivo Excel de Entrada]
          │
          ▼
┌─────────────────────────────────────────┐
│        bpo_decision_making.py           │
│ 1. Lectura y limpieza de datos (Pandas) │
│ 2. Clasificación (Riesgo, Calidad, TMO) │
│ 3. Análisis IA (Google Gemini 2.5 Flash)│
│ 4. Exportación a 3 archivos .xlsx       │
│ 5. Envío opcional de reportes HTML      │
└─────────────────────────────────────────┘
          │
          ▼
   [Carpeta /reports] ──► [config_email.py] ──► [Envío de 3 correos con .xlsx adjuntos]
```

---

## 🛠️ Requisitos Previos e Instalación

* **Python:** Versión `3.9` o superior (se recomienda Python 3.10+).
* **Cuenta de Google AI Studio:** Para obtener una API Key de Gemini (opcional, el sistema funciona sin ella omitiendo el texto de IA).
* **Cuenta de Gmail con Verificación en 2 Pasos:** Necesaria para generar una **Contraseña de Aplicación** y autenticar el envío SMTP.

### Pasos de Instalación

1. **Clonar el repositorio:**
   ```bash
   git clone [https://github.com/carloscastro1298/BPO_Automation.git](https://github.com/carloscastro1298/BPO_Automation.git)
   cd BPO_Automation
   ```

2. **Crear y activar un entorno virtual (recomendado):**
   ```bash
   # En Linux / macOS
   python3 -m venv venv
   source venv/bin/activate

   # En Windows (PowerShell)
   python -m venv venv
   .\venv\Scripts\Activate.ps1
   ```

3. **Instalar las dependencias requeridas:**
   ```bash
   pip install pandas openpyxl google-genai
   ```

---

## ⚙️ Configuración de Variables de Entorno

Antes de ejecutar los scripts, configura las siguientes variables de entorno según las funcionalidades que desees utilizar:

| Variable | Obligatoria | Script que la usa | Descripción | Valor por defecto |
| :--- | :---: | :--- | :--- | :--- |
| `GEMINI_API_KEY` | No | `bpo_decision_making.py` | API Key de Google Gemini para generar análisis automáticos. | `None` (omite IA) |
| `SMTP_USER` | Sí (para correo) | Ambos scripts | Dirección de correo electrónico remitente (ej. Gmail). | `None` |
| `SMTP_PASSWORD` | Sí (para correo) | Ambos scripts | Contraseña de aplicación de 16 caracteres de Gmail. | `None` |
| `BPO_EMAIL_TO` | Sí (para HTML) | `bpo_decision_making.py` | Destinatario para los reportes en formato HTML. | `None` |
| `SMTP_HOST` | No | `bpo_decision_making.py` | Servidor SMTP de salida. | `smtp.gmail.com` |
| `SMTP_PORT` | No | `bpo_decision_making.py` | Puerto del servidor SMTP (STARTTLS). | `587` |

> **Nota sobre `config_email.py`:** Este script utiliza `smtp.gmail.com` en el puerto `587` y tiene configurado por defecto el destinatario en la constante `RECIPIENT_EMAIL` dentro del código. Asegúrate de actualizar esa variable si deseas enviarlo a otra dirección.

### Ejemplos de Configuración

**En Linux / macOS:**
```bash
export GEMINI_API_KEY="AIzaSyTuApiKeyDeGoogleGemini"
export SMTP_USER="tu_correo@gmail.com"
export SMTP_PASSWORD="abcd efgh ijkl mnop"
export BPO_EMAIL_TO="supervisor@empresa.com"
```

**En Windows (PowerShell):**
```powershell
$env:GEMINI_API_KEY="AIzaSyTuApiKeyDeGoogleGemini"
$env:SMTP_USER="tu_correo@gmail.com"
$env:SMTP_PASSWORD="abcd efgh ijkl mnop"
$env:BPO_EMAIL_TO="supervisor@empresa.com"
```

---

## 📄 Estructura del Archivo Excel de Entrada

El archivo `.xlsx` de origen debe contener obligatoriamente las columnas `Decision` y `Result`. Para aprovechar al máximo el análisis de IA y los reportes HTML, se recomienda incluir todas las columnas detalladas a continuación:

| Columna | Requerida | Tipo de Dato | Descripción |
| :--- | :---: | :--- | :--- |
| `Decision` | **Sí** | Texto | Acción tomada por el agente (ej. `Listen first`, `Ask a supervisor`). |
| `Result` | **Sí** | Texto | Desenlace del caso (ej. `Problem solved`, `Case solved`). |
| `Employee_Name` | Recomendada | Texto | Nombre del agente o colaborador que gestionó el caso. |
| `Situation` | Recomendada | Texto | Descripción del problema o solicitud del cliente. |
| `Reason` | Recomendada | Texto | Justificación de la decisión tomada por el agente. |
| `Confidence` | Recomendada | Texto / Num | Nivel de seguridad del agente al tomar la decisión. |
| `Resolution_Time_Minutes` | Recomendada | Numérico | Tiempo total de atención en minutos (valores `>= 0`). |

---

## 💻 Guía de Uso

### 1. Procesamiento Principal y Generación de Reportes (`bpo_decision_making.py`)

Este script admite argumentos por línea de comandos (CLI):

* `input_file`: Ruta al archivo Excel de entrada (obligatorio).
* `--output-dir`: Carpeta donde se guardarán los reportes generados (por defecto: `reports`).
* `--send-email`: Bandera opcional para enviar los resúmenes en HTML por correo electrónico al finalizar el análisis.

**Ejemplos de ejecución:**

```bash
# Procesar casos y generar únicamente los 3 archivos Excel en la carpeta 'reports/'
python bpo_decision_making.py "BPO_Decision_Making_Expanded.xlsx"

# Guardar los reportes en una carpeta personalizada
python bpo_decision_making.py "BPO_Decision_Making_Expanded.xlsx" --output-dir resultados_noviembre

# Procesar casos, generar Excel y enviar los reportes HTML por correo electrónico
python bpo_decision_making.py "BPO_Decision_Making_Expanded.xlsx" --send-email
```

### 2. Envío de Reportes Excel Adjuntos (`config_email.py`)

Una vez generados los tres archivos dentro de la carpeta `reports/`, puedes ejecutar el orquestador de correos para enviar cada archivo `.xlsx` como adjunto con su respectivo asunto y cuerpo institucional:

```bash
python config_email.py
```

Este comando enviará tres correos electrónicos independientes:
1. **ALERTA - Reporte de casos criticos BPO** (`Critical_Cases_Report.xlsx`)
2. **Reporte de casos no criticos BPO** (`Non_Critical_Cases_Report.xlsx`)
3. **Informe general de resultados - BPO Decision Making** (`BPO_Complete_Analysis.xlsx`)

---

## 📊 Reglas de Negocio, Calidad y TMO

### 1. Matriz de Clasificación de Decisiones (`DECISION_RULES`)

| Decisión (`Decision`) | Tipo (`Decision_Type`) | Nivel de Riesgo (`Risk_Level`) | Criticidad (`Criticality`) |
| :--- | :--- | :--- | :--- |
| `Listen first` | Customer Support | `Low` | Non-Critical |
| `Inform the customer` | Communication | `Low` | Non-Critical |
| `Explain the process` | Guidance | `Low` | Non-Critical |
| `Check policy` | Policy | `Medium` | Non-Critical |
| `Investigate the problem` | Investigation | `Medium` | Non-Critical |
| `Ask a supervisor` | Escalation | `High` | **Critical** |
| *Cualquier otro valor* | Other | `Low` | Non-Critical |

### 2. Criterios de Efectividad (`EFFECTIVE_RESULTS`)

La columna `Decision_Quality` se evalúa como **`Effective`** únicamente si el valor de `Result` coincide con los resultados esperados para su nivel de riesgo; de lo contrario, se marca como **`Needs Review`**:

* **Riesgo Bajo (`Low`):** `Problem solved`, `Customer accepted`, `Customer understood`
* **Riesgo Medio (`Medium`):** `Refund approved`, `Problem solved`, `Case solved`
* **Riesgo Alto (`High`):** `Case solved`, `Exception approved`

### 3. Evaluación de TMO y Alertas (`TMO_Status`)

| Rango de `Resolution_Time_Minutes` | Estado (`TMO_Status`) | Alerta (`Alert_Required`) |
| :--- | :--- | :---: |
| Vacío, nulo o inválido | `Not Specified` | `False` |
| `<= 5` minutos | `Within Target` | `False` |
| `> 5` y `<= 30` minutos | `Delayed - Within Limit` | `False` |
| `> 30` minutos | `Critical` | **`True`** |

---

## 📂 Estructura del Proyecto

```text
BPO_Automation/
├── bpo_decision_making.py          # Script principal: ETL, reglas de negocio, Gemini AI, exportación y correo HTML
├── config_email.py                 # Script auxiliar: envío masivo de los reportes Excel como adjuntos vía SMTP
├── BPO_Decision_Making_Expanded.xlsx # Archivo Excel de ejemplo / entrada (opcional)
└── reports/                        # Directorio creado automáticamente al ejecutar el flujo
    ├── Critical_Cases_Report.xlsx
    ├── Non_Critical_Cases_Report.xlsx
    └── BPO_Complete_Analysis.xlsx
```

---

## 🛡️ Manejo de Errores y Solución de Problemas

* **`FileNotFoundError: Excel file not found`:** Verifica que la ruta y el nombre del archivo Excel pasado por consola sean correctos.
* **`ValueError: Missing required Excel columns`:** Asegúrate de que el archivo de entrada contenga exactamente los encabezados `Decision` y `Result`.
* **Análisis de IA omitido:** Si no configuras `GEMINI_API_KEY`, la columna `AI_Analysis` indicará que no se generó el análisis por falta de clave, sin interrumpir la generación de los reportes Excel. Si ocurre un error de red o cuota con la API, el error quedará registrado en esa misma celda sin detener el proceso.
* **Error de autenticación SMTP (`SMTPAuthenticationError`):** Gmail no permite usar la contraseña habitual de la cuenta en scripts externos. Debes activar la verificación en dos pasos en tu cuenta de Google y generar una **Contraseña de Aplicación** específica para `SMTP_PASSWORD`.
* **Error en `config_email.py` (`No se encontro ...`):** Este script valida de forma preventiva que existan los tres archivos `.xlsx` en `reports/` antes de abrir la conexión con Gmail. Ejecuta primero `bpo_decision_making.py` para generarlos.

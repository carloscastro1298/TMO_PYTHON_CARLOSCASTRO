# 📊 BPO Decision Making Automation

![Python](https://img.shields.io/badge/Python-3.8%2B-blue?logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-Data%20Analysis-150458?logo=pandas&logoColor=white)
![Gemini AI](https://img.shields.io/badge/Google%20Gemini-2.5%20Flash-8E75B2?logo=google&logoColor=white)
![Status](https://img.shields.io/badge/Status-Active-success)

> Automated Python system for processing, classifying, AI-analyzing, and reporting BPO operations (replacement for n8n workflows).

---

## 🚀 Key Features

* **📥 Excel Processing:** Reads and cleans data from source Excel files (`.xlsx`).
* **🏷️ Smart Classification:** Automatically categorizes decisions by type, risk level, handling quality, criticality, and AHT (Average Handle Time / TMO) status.
* **🤖 AI Analysis (Google Gemini):** Uses the `gemini-2.5-flash` model to generate concise recommendations and explanations tailored to critical and non-critical cases.
* **📑 Report Generation:** Automatically exports separate Excel reports (`Critical_Cases_Report.xlsx`, `Non_Critical_Cases_Report.xlsx`, `BPO_Complete_Analysis.xlsx`).
* **📧 Email Notifications:** Sends automated emails with HTML summaries and/or Excel report attachments via Gmail's SMTP server.

---

## 🛠️ Requirements & Installation

Make sure you have **Python 3.8 or higher** installed. Then, install the required dependencies by running the following command in your terminal:

```bash
pip install pandas openpyxl google-genai

# 📊 BPO Decision Making Automation

![Python](https://img.shields.io/badge/Python-3.8%2B-blue?logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-Data%20Analysis-150458?logo=pandas&logoColor=white)
![Gemini AI](https://img.shields.io/badge/Google%20Gemini-2.5%20Flash-8E75B2?logo=google&logoColor=white)
![Status](https://img.shields.io/badge/Status-Active-success)

> Sistema automatizado en Python para el procesamiento, clasificación, análisis con IA y generación de reportes de operaciones BPO (reemplazo de flujos de n8n).

---

## 🚀 Características Principales

* **📥 Procesamiento de Excel:** Lee y limpia datos de archivos de origen en formato Excel (`.xlsx`).
* **🏷️ Clasificación Inteligente:** Categoriza automáticamente las decisiones según su tipo, nivel de riesgo, calidad de gestión, criticidad y estado del TMO (Tiempo Medio de Operación).
* **🤖 Análisis con IA (Google Gemini):** Utiliza el modelo `gemini-2.5-flash` para generar recomendaciones y explicaciones en inglés básico adaptadas a casos críticos y no críticos.
* **📑 Generación de Informes:** Exporta automáticamente reportes separados en Excel (`Critical_Cases_Report.xlsx`, `Non_Critical_Cases_Report.xlsx`, `BPO_Complete_Analysis.xlsx`).
* **📧 Notificaciones por Correo:** Envía correos electrónicos automatizados con resúmenes en formato HTML y/o adjuntando los informes en Excel mediante el servidor SMTP de Gmail.

---

## 🛠️ Requisitos e Instalación

Asegúrate de tener instalado **Python 3.8 o superior**. Luego, instala las dependencias necesarias ejecutando el siguiente comando en tu terminal:

```bash
pip install pandas openpyxl google-genai

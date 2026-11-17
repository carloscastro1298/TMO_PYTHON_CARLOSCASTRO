# BPO Decision Making — Versión en Python
# ------------------------------------
# Reemplaza el flujo principal de n8n:
# 1. Lee un archivo de Excel.
# 2. Clasifica cada decisión (tipo, riesgo, calidad, criticidad y TMO).
# 3. Opcionalmente solicita un análisis corto a Google Gemini.
# 4. Exporta informes separando casos críticos y no críticos en Excel.
# 5. Opcionalmente envía resúmenes en HTML por correo electrónico si se configuran los parámetros SMTP.
#
# Instalación:
#     pip install pandas openpyxl google-genai
#
# Ejecución:
#     python bpo_decision_making.py "BPO_Decision_Making_Expanded.xlsx"

from __future__ import annotations

import argparse
import html
import os
import smtplib
from email.message import EmailMessage
from pathlib import Path
from typing import Any

import pandas as pd

# Reglas predefinidas para asociar decisiones con su tipo y nivel de riesgo
DECISION_RULES = {
    "Listen first": ("Customer Support", "Low"),
    "Inform the customer": ("Communication", "Low"),
    "Check policy": ("Policy", "Medium"),
    "Ask a supervisor": ("Escalation", "High"),
    "Explain the process": ("Guidance", "Low"),
    "Investigate the problem": ("Investigation", "Medium"),
}

# Resultados considerados efectivos según el nivel de riesgo
EFFECTIVE_RESULTS = {
    "Low": {"Problem solved", "Customer accepted", "Customer understood"},
    "Medium": {"Refund approved", "Problem solved", "Case solved"},
    "High": {"Case solved", "Exception approved"},
}


def clean_text(value: Any, default: str = "") -> str:
    # Convierte un valor de Excel a texto limpio sin espacios sobrantes,
    # evitando que los valores NaN de Pandas se conviertan en la cadena 'nan'.
    if pd.isna(value):
        return default
    return str(value).strip()


def parse_resolution_time(value: Any) -> float | None:
    # Analiza y retorna un tiempo de resolución válido (no negativo).
    # Retorna None si el valor no está especificado o no es válido.
    if pd.isna(value) or clean_text(value) == "":
        return None
    try:
        number = float(value)
        return number if number >= 0 else None
    except (TypeError, ValueError):
        return None


def classify_case(row: dict[str, Any]) -> dict[str, Any]:
    # Toma una fila de datos (diccionario) de un caso y evalúa su tipo de decisión,
    # nivel de riesgo, calidad, criticidad y estado del TMO (tiempo medio de operación).
    # Retorna el diccionario actualizado con todas las métricas calculadas.
    decision = clean_text(row.get("Decision"))
    result = clean_text(row.get("Result"))

    # Obtener el tipo de decisión y nivel de riesgo basados en las reglas
    decision_type, risk_level = DECISION_RULES.get(decision, ("Other", "Low"))

    # Determinar si la calidad de la decisión fue efectiva o requiere revisión
    decision_quality = (
        "Effective"
        if result in EFFECTIVE_RESULTS.get(risk_level, set())
        else "Needs Review"
    )

    # Definir la criticidad según el nivel de riesgo
    criticality = "Critical" if risk_level == "High" else "Non-Critical"

    # Procesar el tiempo de resolución en minutos
    resolution_time = parse_resolution_time(
        row.get("Resolution_Time_Minutes")
    )

    # Evaluar el estado del TMO (Tiempo Medio de Operación) y si requiere alerta
    if resolution_time is None:
        tmo_status = "Not Specified"
        alert_required = False
    elif resolution_time <= 5:
        tmo_status = "Within Target"
        alert_required = False
    elif resolution_time <= 30:
        tmo_status = "Delayed - Within Limit"
        alert_required = False
    else:
        tmo_status = "Critical"
        alert_required = True

    return {
        **row,
        "Decision_Type": decision_type,
        "Risk_Level": risk_level,
        "Decision_Quality": decision_quality,
        "Criticality": criticality,
        "Resolution_Time_Minutes": resolution_time,
        "TMO_Status": tmo_status,
        "Alert_Required": alert_required,
    }


def gemini_analysis(case: dict[str, Any], api_key: str | None) -> str:
    # Genera un análisis breve del caso utilizando la API de Google Gemini (si está configurada la API Key).
    # Distingue entre casos críticos y no críticos para dar recomendaciones enfocadas.
    if not api_key:
        return (
            "AI analysis was not generated because GEMINI_API_KEY is not set. "
            "Review the situation, decision, result, risk level and TMO status."
        )

    try:
        from google import genai

        client = genai.Client(api_key=api_key)
        critical = case["Criticality"] == "Critical"
        instruction = (
            "Explain why this case is critical and give one clear immediate action."
            if critical
            else "Explain why this case is non-critical and give one practical suggestion."
        )
        prompt = f"""
Analyze this BPO decision using basic English. Keep the response short.

Employee: {clean_text(case.get("Employee_Name"), "Unknown Employee")}
Situation: {clean_text(case.get("Situation"), "Not specified")}
Decision: {clean_text(case.get("Decision"), "Not specified")}
Reason: {clean_text(case.get("Reason"), "Not specified")}
Result: {clean_text(case.get("Result"), "Not specified")}
Confidence: {clean_text(case.get("Confidence"), "Not specified")}
Decision Type: {case["Decision_Type"]}
Risk Level: {case["Risk_Level"]}
Decision Quality: {case["Decision_Quality"]}
Criticality: {case["Criticality"]}
Resolution Time: {case["Resolution_Time_Minutes"]} minutes
TMO Category: {case["TMO_Status"]}

{instruction}
Mention the possible risk or impact, whether the decision was good or needs review,
and whether the resolution time meets the target.
"""
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=prompt,
        )
        return (response.text or "").strip() or "No AI analysis returned."
    except Exception as exc:
        return f"AI analysis unavailable: {exc}"


def analyze_cases(df: pd.DataFrame, api_key: str | None) -> pd.DataFrame:
    # Procesa un DataFrame completo de casos, aplicando la clasificación de cada fila
    # y añadiendo el análisis generado por inteligencia artificial. Retorna un nuevo DataFrame.
    records = df.where(pd.notna(df), None).to_dict(orient="records")
    analyzed = []
    for record in records:
        case = classify_case(record)
        case["AI_Analysis"] = gemini_analysis(case, api_key)
        analyzed.append(case)
    return pd.DataFrame(analyzed)


def build_email_html(cases: pd.DataFrame, title: str, critical: bool) -> str:
    # Construye una plantilla HTML estructurada con el reporte detallado de los casos
    # para ser enviada por correo electrónico. Adapta los colores según la criticidad.
    accent = "#a62e25" if critical else "#176447"
    background = "#fff1ef" if critical else "#e8f6ef"
    case_label = "CRITICAL CASE" if critical else "NON-CRITICAL CASE"

    blocks = []
    for index, row in enumerate(cases.to_dict(orient="records"), start=1):
        employee = html.escape(clean_text(row.get("Employee_Name"), "Unknown Employee"))
        situation = html.escape(clean_text(row.get("Situation"), "Not specified"))
        decision = html.escape(clean_text(row.get("Decision"), "Not specified"))
        analysis = html.escape(clean_text(row.get("AI_Analysis"), "No AI analysis available."))
        resolution = row.get("Resolution_Time_Minutes")
        resolution_text = "Not Specified" if resolution is None or pd.isna(resolution) else f"{resolution} minutes"

        blocks.append(f"""
        <section style="margin:22px 0;border:1px solid #dce3ed;border-radius:10px;overflow:hidden;">
          <header style="background:{background};padding:15px 18px;">
            <small style="color:{accent};font-weight:bold;">{case_label} {index}</small>
            <h2 style="margin:6px 0 0;color:#172b4d;">{employee}</h2>
          </header>
          <div style="padding:18px;color:#526176;line-height:1.6;">
            <p><b>Situation:</b><br>{situation}</p>
            <p><b>Decision:</b><br>{decision}</p>
            <table style="width:100%;border-collapse:collapse;">
              <tr><td style="padding:9px;border:1px solid #e1e7ef;"><b>Criticality</b><br>{html.escape(clean_text(row.get("Criticality")))}</td>
                  <td style="padding:9px;border:1px solid #e1e7ef;"><b>Risk Level</b><br>{html.escape(clean_text(row.get("Risk_Level")))}</td></tr>
              <tr><td style="padding:9px;border:1px solid #e1e7ef;"><b>Decision Quality</b><br>{html.escape(clean_text(row.get("Decision_Quality")))}</td>
                  <td style="padding:9px;border:1px solid #e1e7ef;"><b>Resolution Time</b><br>{resolution_text}</td></tr>
              <tr><td colspan="2" style="padding:9px;border:1px solid #e1e7ef;">
                <b>TMO Category:</b> {html.escape(clean_text(row.get("TMO_Status")))}<br>
                <b>Alert Required:</b> {"Yes" if bool(row.get("Alert_Required")) else "No"}
              </td></tr>
            </table>
            <div style="margin-top:18px;padding:15px;background:#f4f7fb;border-left:4px solid #4263a5;">
              <b>AI Analysis &amp; Recommendation</b><p style="white-space:pre-wrap;">{analysis}</p>
            </div>
          </div>
        </section>
        """)

    return f"""<!doctype html>
    <html><body style="font-family:Arial,sans-serif;background:#f4f7fb;padding:22px;color:#243247;">
      <main style="max-width:760px;margin:auto;background:white;padding:24px;">
        <h1>{html.escape(title)}</h1>
        <p>Review the cases, resolution times and recommendations below.</p>
        {''.join(blocks)}
        <p>This report was generated automatically by BPO Decision Making.</p>
      </main>
    </body></html>"""


def send_report_email(subject: str, body_html: str) -> None:
    # Envía el reporte generado en formato HTML mediante el servidor SMTP configurado
    # en las variables de entorno.
    recipient = os.getenv("BPO_EMAIL_TO")
    sender = os.getenv("SMTP_USER")
    password = os.getenv("SMTP_PASSWORD")
    if not all((recipient, sender, password)):
        print(f"Email skipped: configure BPO_EMAIL_TO, SMTP_USER and SMTP_PASSWORD ({subject}).")
        return

    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = sender
    message["To"] = recipient
    message.set_content("This report is available in HTML format.")
    message.add_alternative(body_html, subtype="html")

    host = os.getenv("SMTP_HOST", "smtp.gmail.com")
    port = int(os.getenv("SMTP_PORT", "587"))
    with smtplib.SMTP(host, port) as server:
        server.starttls()
        server.login(sender, password)
        server.send_message(message)
    print(f"Email sent: {subject}")


def main() -> None:
    # Función principal de ejecución por línea de comandos (CLI).
    # Lee el archivo de Excel de entrada, procesa el análisis, exporta los reportes
    # en la carpeta de salida y opcionalmente dispara los correos electrónicos.
    parser = argparse.ArgumentParser(description="BPO Decision Making workflow in Python")
    parser.add_argument("input_file", help="Path to the source Excel file")
    parser.add_argument("--output-dir", default="reports", help="Folder for generated reports")
    parser.add_argument("--send-email", action="store_true", help="Send reports if SMTP settings exist")
    args = parser.parse_args()

    input_path = Path(args.input_file)
    if not input_path.exists():
        raise FileNotFoundError(f"Excel file not found: {input_path}")

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    source_df = pd.read_excel(input_path)
    required_columns = {"Decision", "Result"}
    missing = required_columns - set(source_df.columns)
    if missing:
        raise ValueError(f"Missing required Excel columns: {', '.join(sorted(missing))}")

    api_key = os.getenv("GEMINI_API_KEY")
    results = analyze_cases(source_df, api_key)

    critical_cases = results[results["Criticality"] == "Critical"].copy()
    non_critical_cases = results[results["Criticality"] == "Non-Critical"].copy()

    critical_path = output_dir / "Critical_Cases_Report.xlsx"
    non_critical_path = output_dir / "Non_Critical_Cases_Report.xlsx"
    complete_path = output_dir / "BPO_Complete_Analysis.xlsx"

    critical_cases.to_excel(critical_path, index=False)
    non_critical_cases.to_excel(non_critical_path, index=False)
    results.to_excel(complete_path, index=False)

    print(f"Cases processed: {len(results)}")
    print(f"Critical cases: {len(critical_cases)}")
    print(f"Non-critical cases: {len(non_critical_cases)}")
    print(f"Reports saved in: {output_dir.resolve()}")

    if args.send_email:
        if not critical_cases.empty:
            send_report_email(
                "Critical BPO Cases - Action Required",
                build_email_html(critical_cases, "Critical Cases Report", True),
            )
        if not non_critical_cases.empty:
            send_report_email(
                "BPO Non-Critical Cases - Analysis Report",
                build_email_html(non_critical_cases, "Non-Critical Cases Report", False),
            )


if __name__ == "__main__":
    main()
import os
import smtplib
from pathlib import Path
from email.message import EmailMessage

# CONFIGURACIÓN DE GMAIL (Servidor y Puerto SMTP predeterminados)
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# Obtención de credenciales de remitente desde las variables de entorno del sistema
SENDER_EMAIL = os.getenv("SMTP_USER")
SENDER_PASSWORD = os.getenv("SMTP_PASSWORD")

# Correo electrónico predeterminado del destinatario
RECIPIENT_EMAIL = "cdetomartinez1209@gmail.com"

# Ruta local de la carpeta donde se almacenan los reportes generados
CARPETA_REPORTES = Path(__file__).resolve().parent / "reports"


def crear_correo(asunto, mensaje, archivo):
    # Construye un objeto EmailMessage con el asunto, remitente, destinatario,
    # cuerpo del mensaje en texto y adjunta un archivo Excel de reportes.
    correo = EmailMessage()

    correo["Subject"] = asunto
    correo["From"] = SENDER_EMAIL
    correo["To"] = RECIPIENT_EMAIL

    correo.set_content(mensaje)

    # Lectura y adjunto del archivo Excel especificado
    with open(archivo, "rb") as documento:
        correo.add_attachment(
            documento.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=archivo.name
        )

    return correo


def enviar_informes():
    # Función principal para orquestar el envío masivo de los tres informes de BPO
    # (Casos Críticos, Casos No Críticos y Análisis General) conectándose al servidor SMTP de Gmail.
    print("\nINICIANDO SISTEMA BPO DECISION MAKING")

    # Validación previa de las credenciales de correo
    if not SENDER_EMAIL or not SENDER_PASSWORD:
        print("ERROR: Faltan las credenciales de Gmail.")
        return

    # Lista estructurada con los reportes a enviar, su asunto y mensaje personalizado
    reportes = [
        {
            "archivo": "Critical_Cases_Report.xlsx",
            "asunto": "ALERTA - Reporte de casos criticos BPO",
            "mensaje": """
Cordial saludo,

Adjuntamos el reporte de casos criticos generado
por el sistema BPO Decision Making.

Este informe contiene las situaciones que requieren
atencion prioritaria y seguimiento.

Recomendamos revisar los resultados para identificar
oportunidades de mejora en la gestion y resolucion
de los casos.

Atentamente,

Departamento de Analisis y Calidad
BPO Decision Making
"""
        },
        {
            "archivo": "Non_Critical_Cases_Report.xlsx",
            "asunto": "Reporte de casos no criticos BPO",
            "mensaje": """
Cordial saludo,

Compartimos el reporte de casos no criticos generado
por el sistema BPO Decision Making.

Este documento permite realizar seguimiento a los
casos gestionados dentro de los niveles normales
de operacion.

Agradecemos su revision.

Atentamente,

Departamento de Analisis y Calidad
BPO Decision Making
"""
        },
        {
            "archivo": "BPO_Complete_Analysis.xlsx",
            "asunto": "Informe general de resultados - BPO Decision Making",
            "mensaje": """
Cordial saludo,

Compartimos el informe general del sistema
BPO Decision Making.

El documento contiene el analisis completo de los
casos procesados, su clasificacion, los tiempos
de resolucion y los resultados obtenidos.

Este analisis busca fortalecer la calidad del
servicio, la toma de decisiones y la atencion
al cliente.

Atentamente,

Departamento de Analisis y Calidad
BPO Decision Making
"""
        }
    ]

    # Verificar de manera preventiva que existan todos los archivos de reportes en disco
    for reporte in reportes:
        ruta = CARPETA_REPORTES / reporte["archivo"]

        if not ruta.exists():
            print(f"ERROR: No se encontro {ruta}")
            return

    print("Conectando con Gmail...")

    try:
        # Establecer conexión segura por SMTP con STARTTLS
        with smtplib.SMTP(
            SMTP_SERVER, SMTP_PORT, timeout=30
        ) as servidor:

            servidor.starttls()

            # Autenticación con las credenciales (removiendo espacios en la contraseña si los hay)
            servidor.login(
                SENDER_EMAIL,
                SENDER_PASSWORD.replace(" ", "")
            )

            print("Conexion con Gmail exitosa.")

            # Iterar sobre cada reporte para armar y enviar su respectivo correo
            for reporte in reportes:

                ruta = CARPETA_REPORTES / reporte["archivo"]

                correo = crear_correo(
                    reporte["asunto"],
                    reporte["mensaje"],
                    ruta
                )

                servidor.send_message(correo)

                print(
                    "Correo enviado:",
                    reporte["asunto"]
                )

        print("\nLOS TRES CORREOS FUERON ENVIADOS CORRECTAMENTE.")

    except Exception as error:
        print("\nERROR AL ENVIAR LOS CORREOS:", error)


if __name__ == "__main__":
    enviar_informes()
import pandas as pd
import requests
import io
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dateutil.relativedelta import relativedelta

# =====================================================
# CONFIGURACION
# =====================================================

EXCEL_URL = "https://valserindustriales-my.sharepoint.com/personal/sst_valserindustriales_com/_layouts/15/download.aspx?share=IQDAwbM-LAyqSYzysaGTnRooAUZJRuK8wwFzV7eMWEMJ-DU"

CORREOS_DESTINO = [
    "Tecnicodeservicios@valserindustriales.com"
]

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASS = os.getenv("SMTP_PASS")

# Dias antes del vencimiento para alertar
DIAS_ALERTA = 30

# =====================================================
# DESCARGAR EXCEL DESDE ONEDRIVE
# =====================================================

def descargar_excel():

    headers = {
        "User-Agent": "Mozilla/5.0"
    }

    r = requests.get(
        EXCEL_URL,
        headers=headers,
        allow_redirects=True
    )

    if r.status_code != 200:
        raise Exception(
            f"No fue posible descargar el Excel. HTTP {r.status_code}"
        )

    if len(r.content) < 1000:
        raise Exception(
            "El archivo descargado parece inválido o vacío."
        )

    return r.content

# =====================================================
# CARGAR EXCEL
# =====================================================

def cargar_excel(bytes_excel):

    df = pd.read_excel(
        io.BytesIO(bytes_excel),
        sheet_name="LISTADO MAESTRO",
        header=3
    )

    # Limpiar nombres de columnas
    df.columns = [str(c).strip() for c in df.columns]

    print("COLUMNAS ENCONTRADAS:")
    for c in df.columns:
        print(repr(c))

    # Eliminar filas completamente vacías
    df = df.dropna(how="all")

    return df

# =====================================================
# PREPARAR DATOS
# =====================================================

def preparar_datos(df):

    hoy = pd.Timestamp.now().normalize()

    # Convertir fecha
    df["FECHA"] = pd.to_datetime(
        df["FECHA"],
        errors="coerce"
    )

    # Limpiar periodicidad
    df["PERIODICIDAD REVISIÓN"] = (
        df["PERIODICIDAD REVISIÓN"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    fechas_vencimiento = []

    for _, row in df.iterrows():

        fecha_revision = row["FECHA"]
        periodicidad = row["PERIODICIDAD REVISIÓN"]

        if pd.isna(fecha_revision):
            fechas_vencimiento.append(pd.NaT)
            continue

        if periodicidad == "ANUAL":
            vencimiento = fecha_revision + relativedelta(years=1)

        elif periodicidad == "BIENAL":
            vencimiento = fecha_revision + relativedelta(years=2)

        else:
            vencimiento = pd.NaT

        fechas_vencimiento.append(vencimiento)

    df["FECHA_VENCIMIENTO"] = fechas_vencimiento

    df["DIAS_RESTANTES"] = (
        df["FECHA_VENCIMIENTO"] - hoy
    ).dt.days

    return df

# =====================================================
# CLASIFICAR ALERTAS
# =====================================================

def clasificar_alertas(df):

    # Próximos a vencer
    proximos = df[
        (df["DIAS_RESTANTES"] >= 0) &
        (df["DIAS_RESTANTES"] <= DIAS_ALERTA)
    ]

    # Vencidos
    vencidos = df[
        (df["DIAS_RESTANTES"] < 0)
    ]

    return proximos, vencidos

# =====================================================
# GENERAR TABLA HTML
# =====================================================

def generar_tabla(df, titulo):

    if df.empty:
        return ""

    filas = ""

    for _, row in df.iterrows():

        proceso = (
            str(row["PROCESO"])
            .replace("\n", "")
            .replace("\r", "")
            .strip()
        )

        tipo = (
            str(row["TIPO"])
            .replace("\n", "")
            .replace("\r", "")
            .strip()
        )

        consecutivo = (
            str(row["CONSECUTIVO"])
            .replace("\n", "")
            .replace("\r", "")
            .strip()
        )

        # Convertir 1 -> 01
        if consecutivo.isdigit():
            consecutivo = consecutivo.zfill(2)

        codigo = f"{proceso}-{tipo}-{consecutivo}"

        # Limpieza final
        codigo = " ".join(codigo.split())

        nombre = str(row["NOMBRE DEL DOCUMENTO"]).strip()

        fecha_vencimiento = (
            row["FECHA_VENCIMIENTO"].strftime("%Y-%m-%d")
            if pd.notna(row["FECHA_VENCIMIENTO"])
            else "Sin fecha"
        )

        dias = row["DIAS_RESTANTES"]

        filas += f"""
        <tr>
            <td style="white-space: nowrap;">{codigo}</td>
            <td>{nombre}</td>
            <td>{fecha_vencimiento}</td>
            <td>{dias}</td>
        </tr>
        """

    html = f"""
    <h2>{titulo}</h2>

    <table border="1" cellspacing="0" cellpadding="5" style="border-collapse: collapse; font-family: Arial;">
        <tr style="background-color: #f2f2f2;">
            <th>CODIGO</th>
            <th>NOMBRE</th>
            <th>FECHA VENCIMIENTO</th>
            <th>DIAS</th>
        </tr>

        {filas}

    </table>

    <br>
    """

    return html

# =====================================================
# ENVIAR CORREO
# =====================================================

def enviar_correo(asunto, html):

    msg = MIMEMultipart()

    msg["From"] = SMTP_USER
    msg["To"] = ", ".join(CORREOS_DESTINO)
    msg["Subject"] = asunto

    msg.attach(MIMEText(html, "html"))

    server = smtplib.SMTP(
        SMTP_SERVER,
        SMTP_PORT
    )

    server.starttls()

    server.login(
        SMTP_USER,
        SMTP_PASS
    )

    server.sendmail(
        SMTP_USER,
        CORREOS_DESTINO,
        msg.as_string()
    )

    server.quit()

# =====================================================
# MAIN
# =====================================================

def main():

    print("Descargando Excel...")

    excel_bytes = descargar_excel()

    print("Leyendo Excel...")

    df = cargar_excel(excel_bytes)

    print("Preparando datos...")

    print(df.columns.tolist())

    df = preparar_datos(df)

    proximos, vencidos = clasificar_alertas(df)

    print(f"Proximos a vencer: {len(proximos)}")
    print(f"Vencidos: {len(vencidos)}")

    if proximos.empty and vencidos.empty:
        print("No hay alertas para enviar")
        return

    html = "<h1>Alerta revisión documental OHSQ-FO-34</h1>"

    html += generar_tabla(
        proximos,
        "Documentos próximos a vencer"
    )

    html += generar_tabla(
        vencidos,
        "Documentos vencidos"
    )

    enviar_correo(
        "Alerta revisión documental OHSQ-FO-34",
        html
    )

    print("Correo enviado correctamente")

# =====================================================
# EJECUCION
# =====================================================

if __name__ == "__main__":
    main()

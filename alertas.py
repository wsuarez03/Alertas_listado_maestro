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
    "Tecnicodeservicios@valserindustriales.com",
   # "sst@valserindustriales.com",
   # "asesorcomercial@valserindustriales.com",
   # "proyectos@valserindustriales.com",
   # "contabilidad@valserindustriales.com",
   # "comercial@valserindustriales.com"
    
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

    html = f"<h2>{titulo}</h2>"

    # ============================================
    # LIMPIAR COLUMNAS
    # ============================================

    df["PROCESO_LIMPIO"] = (
        df["PROCESO"]
        .astype(str)
        .str.replace("\n", "", regex=False)
        .str.replace("\r", "", regex=False)
        .str.strip()
    )

    df["TIPO_LIMPIO"] = (
        df["TIPO"]
        .astype(str)
        .str.replace("\n", "", regex=False)
        .str.replace("\r", "", regex=False)
        .str.strip()
    )

    df["CONSECUTIVO_LIMPIO"] = (
        df["CONSECUTIVO"]
        .astype(str)
        .str.replace("\n", "", regex=False)
        .str.replace("\r", "", regex=False)
        .str.strip()
    )

    # ============================================
    # CONSECUTIVO NUMERICO
    # ============================================

    df["CONSECUTIVO_ORDEN"] = pd.to_numeric(
        df["CONSECUTIVO_LIMPIO"],
        errors="coerce"
    )

    # ============================================
    # CLASIFICAR SEVERIDAD
    # ============================================

    severidades = []
    colores = []
    orden_criticidad = []

    for dias in df["DIAS_RESTANTES"]:

        # PROXIMOS A VENCER
        if dias >= 0:

            if dias <= 7:
                severidades.append(f"🟠 Vence en {int(dias)} días")
                colores.append("#fff3cd")
                orden_criticidad.append(5)

            else:
                severidades.append(f"🟡 Vence en {int(dias)} días")
                colores.append("#fff8cc")
                orden_criticidad.append(6)

        # VENCIDOS
        else:

            vencido = abs(int(dias))

            if vencido > 365:
                severidades.append(f"🚨 Vencido hace {vencido} días")
                colores.append("#f8d7da")
                orden_criticidad.append(1)

            elif vencido > 180:
                severidades.append(f"🔴 Vencido hace {vencido} días")
                colores.append("#f5c6cb")
                orden_criticidad.append(2)

            elif vencido > 30:
                severidades.append(f"🟠 Vencido hace {vencido} días")
                colores.append("#ffe5b4")
                orden_criticidad.append(3)

            else:
                severidades.append(f"🟡 Vencido hace {vencido} días")
                colores.append("#fff3cd")
                orden_criticidad.append(4)

    df["ESTADO"] = severidades
    df["COLOR"] = colores
    df["ORDEN_CRITICIDAD"] = orden_criticidad

    # ============================================
    # RESUMEN EJECUTIVO
    # ============================================

    total = len(df)

    criticos = len(df[df["ORDEN_CRITICIDAD"] == 1])

    procesos_mayores = (
        df.groupby("PROCESO_LIMPIO")
        .size()
        .sort_values(ascending=False)
        .head(5)
    )

    resumen_procesos = ""

    for proceso, cantidad in procesos_mayores.items():

        resumen_procesos += f"""
        <li>
            <b>{proceso}</b>: {cantidad}
        </li>
        """

    html += f"""
    <div style="
        background-color:#f4f4f4;
        padding:15px;
        border-radius:8px;
        margin-bottom:25px;
        font-family:Arial;
    ">

        <h3 style="margin-top:0;">
            📌 Resumen Ejecutivo
        </h3>

        <ul>
            <li><b>Total documentos:</b> {total}</li>
            <li><b>Documentos críticos:</b> {criticos}</li>
        </ul>

        <b>Procesos con más registros:</b>

        <ul>
            {resumen_procesos}
        </ul>

    </div>
    """

    # ============================================
    # AGRUPAR POR PROCESO
    # ============================================

    procesos = df.groupby("PROCESO_LIMPIO")

    for proceso, df_proceso in procesos:

        # ============================================
        # ORDENAR POR CRITICIDAD
        # ============================================

        df_proceso = df_proceso.sort_values(
            by=[
                "ORDEN_CRITICIDAD",
                "DIAS_RESTANTES",
                "TIPO_LIMPIO",
                "CONSECUTIVO_ORDEN"
            ]
        )

        html += f"""
        <h2 style="
            background-color:#d9ead3;
            padding:10px;
            border-radius:6px;
            margin-top:35px;
            font-family:Arial;
        ">
            PROCESO: {proceso}
        </h2>
        """

        filas = ""

        for _, row in df_proceso.iterrows():

            tipo = row["TIPO_LIMPIO"]

            consecutivo = row["CONSECUTIVO_LIMPIO"]

            if consecutivo.isdigit():
                consecutivo = consecutivo.zfill(2)

            codigo = f"{proceso}-{tipo}-{consecutivo}"

            codigo = " ".join(codigo.split())

            fecha_vencimiento = (
                row["FECHA_VENCIMIENTO"].strftime("%Y-%m-%d")
                if pd.notna(row["FECHA_VENCIMIENTO"])
                else "Sin fecha"
            )

            estado = row["ESTADO"]

            color = row["COLOR"]

            filas += f"""
            <tr style="background-color:{color};">

                <td style="
                    white-space: nowrap;
                    font-weight:bold;
                ">
                    {codigo}
                </td>

                <td>
                    {fecha_vencimiento}
                </td>

                <td>
                    {estado}
                </td>

            </tr>
            """

        html += f"""
        <table
            border="1"
            cellspacing="0"
            cellpadding="6"
            style="
                border-collapse: collapse;
                font-family: Arial;
                margin-bottom: 30px;
                width: 100%;
            "
        >

            <tr style="
                background-color:#2f5597;
                color:white;
            ">
                <th>CODIGO</th>
                <th>FECHA VENCIMIENTO</th>
                <th>ESTADO</th>
            </tr>

            {filas}

        </table>
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

    html = """
    <h1 style="
        font-family:Arial;
        background-color:#1f4e78;
        color:white;
        padding:15px;
        border-radius:8px;
    ">
        📋 Alerta revisión documental OHSQ-FO-34
    </h1>
    """

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

```python
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
]

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASS = os.getenv("SMTP_PASS")

DIAS_ALERTA = 30

# =====================================================
# DESCARGAR EXCEL
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

    df.columns = [str(c).strip() for c in df.columns]

    print("COLUMNAS ENCONTRADAS:")
    for c in df.columns:
        print(repr(c))

    df = df.dropna(how="all")

    return df

# =====================================================
# PREPARAR DATOS
# =====================================================

def preparar_datos(df):

    hoy = pd.Timestamp.now().normalize()

    # ============================================
    # FECHAS
    # ============================================

    df["FECHA"] = pd.to_datetime(
        df["FECHA"],
        errors="coerce"
    )

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
    # NOMBRE DOCUMENTO
    # ============================================

    # CAMBIAR "NOMBRE" SI TU COLUMNA SE LLAMA DIFERENTE

    df["NOMBRE_DOCUMENTO"] = (
        df["NOMBRE"]
        .astype(str)
        .str.replace("\n", "", regex=False)
        .str.replace("\r", "", regex=False)
        .str.strip()
    )

    # ============================================
    # PERIODICIDAD
    # ============================================

    df["PERIODICIDAD REVISIÓN"] = (
        df["PERIODICIDAD REVISIÓN"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # ============================================
    # FECHA VENCIMIENTO
    # ============================================

    fechas_vencimiento = []

    for _, row in df.iterrows():

        fecha_revision = row["FECHA"]
        periodicidad = row["PERIODICIDAD REVISIÓN"]

        if pd.isna(fecha_revision):
            fechas_vencimiento.append(pd.NaT)
            continue

        if periodicidad == "ANUAL":

            vencimiento = (
                fecha_revision +
                relativedelta(years=1)
            )

        elif periodicidad == "BIENAL":

            vencimiento = (
                fecha_revision +
                relativedelta(years=2)
            )

        else:
            vencimiento = pd.NaT

        fechas_vencimiento.append(vencimiento)

    df["FECHA_VENCIMIENTO"] = fechas_vencimiento

    # ============================================
    # DIAS RESTANTES
    # ============================================

    df["DIAS_RESTANTES"] = (
        df["FECHA_VENCIMIENTO"] - hoy
    ).dt.days

    # ============================================
    # CONSECUTIVO NUMERICO
    # ============================================

    df["CONSECUTIVO_ORDEN"] = pd.to_numeric(
        df["CONSECUTIVO_LIMPIO"],
        errors="coerce"
    )

    return df

# =====================================================
# CLASIFICAR ALERTAS
# =====================================================

def clasificar_alertas(df):

    proximos = df[
        (df["DIAS_RESTANTES"] >= 0) &
        (df["DIAS_RESTANTES"] <= DIAS_ALERTA)
    ]

    vencidos = df[
        (df["DIAS_RESTANTES"] < 0)
    ]

    return proximos, vencidos

# =====================================================
# CLASIFICAR SEVERIDAD
# =====================================================

def clasificar_severidad(df):

    severidades = []
    colores = []
    ordenes = []

    for dias in df["DIAS_RESTANTES"]:

        # ============================================
        # PROXIMOS
        # ============================================

        if dias >= 0:

            if dias <= 7:

                severidades.append(
                    f"🟠 Vence en {int(dias)} días"
                )

                colores.append("#fff3cd")
                ordenes.append(5)

            else:

                severidades.append(
                    f"🟡 Vence en {int(dias)} días"
                )

                colores.append("#fff8cc")
                ordenes.append(6)

        # ============================================
        # VENCIDOS
        # ============================================

        else:

            vencido = abs(int(dias))

            if vencido > 365:

                severidades.append(
                    f"🚨 Vencido hace {vencido} días"
                )

                colores.append("#f8d7da")
                ordenes.append(1)

            elif vencido > 180:

                severidades.append(
                    f"🔴 Vencido hace {vencido} días"
                )

                colores.append("#f5c6cb")
                ordenes.append(2)

            elif vencido > 30:

                severidades.append(
                    f"🟠 Vencido hace {vencido} días"
                )

                colores.append("#ffe5b4")
                ordenes.append(3)

            else:

                severidades.append(
                    f"🟡 Vencido hace {vencido} días"
                )

                colores.append("#fff3cd")
                ordenes.append(4)

    df["ESTADO"] = severidades
    df["COLOR"] = colores
    df["ORDEN_CRITICIDAD"] = ordenes

    return df

# =====================================================
# DASHBOARD
# =====================================================

def generar_dashboard(df):

    total = len(df)

    criticos = len(
        df[df["ORDEN_CRITICIDAD"] == 1]
    )

    vencidos = len(
        df[df["DIAS_RESTANTES"] < 0]
    )

    proximos = len(
        df[
            (df["DIAS_RESTANTES"] >= 0) &
            (df["DIAS_RESTANTES"] <= DIAS_ALERTA)
        ]
    )

    html = f"""
    <div style="
        display:flex;
        gap:15px;
        flex-wrap:wrap;
        margin-bottom:30px;
        font-family:Arial;
    ">

        <div style="
            flex:1;
            min-width:180px;
            background:#b71c1c;
            color:white;
            padding:20px;
            border-radius:10px;
            text-align:center;
        ">
            <div style="
                font-size:32px;
                font-weight:bold;
            ">
                {criticos}
            </div>

            <div>
                🚨 Críticos
            </div>
        </div>

        <div style="
            flex:1;
            min-width:180px;
            background:#d84315;
            color:white;
            padding:20px;
            border-radius:10px;
            text-align:center;
        ">
            <div style="
                font-size:32px;
                font-weight:bold;
            ">
                {vencidos}
            </div>

            <div>
                🔴 Vencidos
            </div>
        </div>

        <div style="
            flex:1;
            min-width:180px;
            background:#f9a825;
            color:white;
            padding:20px;
            border-radius:10px;
            text-align:center;
        ">
            <div style="
                font-size:32px;
                font-weight:bold;
            ">
                {proximos}
            </div>

            <div>
                🟡 Próximos
            </div>
        </div>

        <div style="
            flex:1;
            min-width:180px;
            background:#1565c0;
            color:white;
            padding:20px;
            border-radius:10px;
            text-align:center;
        ">
            <div style="
                font-size:32px;
                font-weight:bold;
            ">
                {total}
            </div>

            <div>
                📄 Total
            </div>
        </div>

    </div>
    """

    return html

# =====================================================
# GENERAR TABLA
# =====================================================

def generar_tabla(df, titulo):

    if df.empty:
        return ""

    html = f"""
    <h2 style="
        font-family:Arial;
        color:#1f4e78;
        margin-top:40px;
    ">
        {titulo}
    </h2>
    """

    # ============================================
    # AGRUPAR POR PROCESO
    # ============================================

    procesos = df.groupby("PROCESO_LIMPIO")

    for proceso, df_proceso in procesos:

        # ============================================
        # ORDENAR
        # ============================================

        df_proceso = df_proceso.sort_values(
            by=[
                "TIPO_LIMPIO",
                "CONSECUTIVO_ORDEN"
            ]
        )

        html += f"""
        <h3 style="
            background:#d9ead3;
            padding:10px;
            border-radius:6px;
            font-family:Arial;
            margin-top:30px;
        ">
            📂 PROCESO: {proceso}
        </h3>
        """

        filas = ""

        for _, row in df_proceso.iterrows():

            tipo = row["TIPO_LIMPIO"]

            consecutivo = row["CONSECUTIVO_LIMPIO"]

            nombre_documento = row["NOMBRE_DOCUMENTO"]

            if consecutivo.isdigit():
                consecutivo = consecutivo.zfill(2)

            fecha_vencimiento = (
                row["FECHA_VENCIMIENTO"].strftime("%Y-%m-%d")
                if pd.notna(row["FECHA_VENCIMIENTO"])
                else "Sin fecha"
            )

            estado = row["ESTADO"]

            color = row["COLOR"]

            filas += f"""
            <tr style="
                background-color:{color};
            ">

                <td style="
                    font-weight:bold;
                    white-space:nowrap;
                    text-align:center;
                ">
                    {tipo}
                </td>

                <td style="
                    text-align:center;
                    white-space:nowrap;
                ">
                    {consecutivo}
                </td>

                <td style="
                    word-break:break-word;
                ">
                    {nombre_documento}
                </td>

                <td style="
                    white-space:nowrap;
                    text-align:center;
                ">
                    {fecha_vencimiento}
                </td>

                <td style="
                    font-weight:bold;
                    white-space:nowrap;
                ">
                    {estado}
                </td>

            </tr>
            """

        html += f"""
        <table
            border="1"
            cellspacing="0"
            cellpadding="8"
            style="
                border-collapse:collapse;
                font-family:Arial;
                margin-bottom:30px;
                width:100%;
                table-layout:fixed;
                font-size:14px;
            "
        >

            <tr style="
                background-color:#1f4e78;
                color:white;
            ">
                <th style="width:90px;">
                    TIPO
                </th>

                <th style="width:80px;">
                    CONS
                </th>

                <th>
                    NOMBRE DOCUMENTO
                </th>

                <th style="width:140px;">
                    FECHA VENCIMIENTO
                </th>

                <th style="width:220px;">
                    ESTADO
                </th>
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

    df = preparar_datos(df)

    df = clasificar_severidad(df)

    proximos, vencidos = clasificar_alertas(df)

    print(f"Proximos a vencer: {len(proximos)}")
    print(f"Vencidos: {len(vencidos)}")

    if proximos.empty and vencidos.empty:

        print("No hay alertas para enviar")
        return

    total_criticos = len(
        df[df["ORDEN_CRITICIDAD"] == 1]
    )

    asunto = (
        f"🚨 {len(vencidos)} vencidos | "
        f"{total_criticos} críticos | "
        f"OHSQ-FO-34"
    )

    html = f"""
    <div style="
        font-family:Arial;
    ">

        <h1 style="
            background-color:#1f4e78;
            color:white;
            padding:18px;
            border-radius:8px;
        ">
            📋 Alerta revisión documental OHSQ-FO-34
        </h1>

        <p style="
            font-size:15px;
            margin-top:20px;
            margin-bottom:25px;
        ">
            Se requiere priorizar la gestión de los documentos
            marcados como críticos para evitar incumplimientos
            operacionales y de calidad.
        </p>

    </div>
    """

    html += generar_dashboard(df)

    html += generar_tabla(
        proximos,
        "🟡 Documentos próximos a vencer"
    )

    html += generar_tabla(
        vencidos,
        "🔴 Documentos vencidos"
    )

    enviar_correo(
        asunto,
        html
    )

    print("Correo enviado correctamente")

# =====================================================
# EJECUCION
# =====================================================

if __name__ == "__main__":
    main()
```

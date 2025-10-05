# simulador_compuesto.py
# Simulador Web para ABECOIN (Interés Compuesto — cuota fija, amortización variable)
# Requisitos: streamlit, pandas, openpyxl
# Ejecutar: streamlit run simulador_compuesto.py

import streamlit as st
import pandas as pd
import math
from datetime import datetime, timedelta
from io import BytesIO

# --------------------------
# CONFIG
# --------------------------
st.set_page_config(page_title="ABECOIN - Simulador Compuesto", page_icon="🐝", layout="wide")

LOGO_PATH = "abecoin_logo.png"  # Cambia si tu logo tiene otro nombre

AZUL = "#062a6f"
AMARILLO = "#FFD166"
TEXTO_HEADER = "#FFFFFF"

# --------------------------
# FUNCIONES DE TASAS / DEGRAVAMEN
# --------------------------
def obtener_tasa_semanal(capital, cuotas):
    opciones = [2, 3, 4]
    cuotas_closest = min(opciones, key=lambda x: abs(x - cuotas))
    if 10 < capital <= 200:
        if cuotas_closest == 2: return 0.04
        if cuotas_closest == 3: return 0.03
        if cuotas_closest == 4: return 0.025
    elif 200 < capital <= 400:
        if cuotas_closest == 2: return 0.02
        if cuotas_closest == 3: return 0.0167
        if cuotas_closest == 4: return 0.015
    elif 400 < capital <= 600:
        if cuotas_closest == 2: return 0.025
        if cuotas_closest == 3: return 0.02
        if cuotas_closest == 4: return 0.0175
    return 0.03

def obtener_porcentaje_degravamen(capital):
    if capital <= 200:
        return 0.008
    elif capital <= 400:
        return 0.01
    else:
        return 0.015

# --------------------------
# LÓGICA DEL CRONOGRAMA (INTERÉS COMPUESTO — CUOTA FIJA)
# --------------------------
def generar_cronograma(nombre, dni, direccion, capital, cuotas, degrav_mode="prorated"):
    tasa = obtener_tasa_semanal(capital, cuotas)

    # Si la tasa es 0 (caso borde), cuota = capital / cuotas
    if tasa == 0:
        cuota_fija = round(capital / cuotas, 2)
    else:
        cuota_fija = round((capital * tasa) / (1 - (1 + tasa) ** (-cuotas)), 2)

    pct_degrav = obtener_porcentaje_degravamen(capital)
    degrav_total = round(capital * pct_degrav, 2)

    # Degravamen prorrateado o upfront
    if degrav_mode == "prorated":
        base_prorr = math.floor((degrav_total / cuotas) * 100) / 100
        prorrateos = [base_prorr] * cuotas
        diff = round(degrav_total - sum(prorrateos), 2)
        prorrateos[-1] = round(prorrateos[-1] + diff, 2)
    else:
        prorrateos = [0.0] * cuotas
        if cuotas >= 1:
            prorrateos[0] = degrav_total

    hoy = datetime.today()
    saldo = capital
    filas = []
    interes_total = 0.0
    suma_cuotas_base = 0.0
    suma_cuotas_final = 0.0

    for i in range(1, cuotas + 1):
        prev_saldo = saldo
        interes_periodo = round(prev_saldo * tasa, 2)
        # amortización: cuota fija menos interés de este periodo
        amortizacion = round(cuota_fija - interes_periodo, 2)

        # Ajuste en la última cuota para eliminar residuos por redondeo (saldo a 0)
        if i == cuotas:
            # si queda una pequeña diferencia por redondeo, tomamos todo el saldo
            amortizacion = round(prev_saldo, 2)
            # recalcular cuota_base como suma de amortización + interés
            cuota_base = round(amortizacion + interes_periodo, 2)
        else:
            cuota_base = round(cuota_fija, 2)

        cuota_final = round(cuota_base + prorrateos[i - 1], 2)
        vencimiento = (hoy + timedelta(weeks=i)).strftime("%d/%m/%Y")
        estado = "PENDIENTE"

        filas.append({
            "N° Cuota": i,
            "Fecha Venc.": vencimiento,
            "Saldo Capital": round(prev_saldo, 2),   # saldo al inicio del periodo
            "Amortización": amortizacion,
            "Interés": interes_periodo,
            "Cuota Base": cuota_base,
            "Degravamen": prorrateos[i - 1],
            "Cuota Final": cuota_final,
            "Estado": estado
        })

        interes_total += interes_periodo
        suma_cuotas_base += cuota_base
        suma_cuotas_final += cuota_final

        saldo = round(prev_saldo - amortizacion, 2)
        # Evitar -0.0 o pequeños negativos por redondeo
        if abs(saldo) < 0.01:
            saldo = 0.0

    df = pd.DataFrame(filas)

    resumen = {
        "Nombre": nombre,
        "DNI": dni,
        "Dirección": direccion,
        "Capital Inicial (S/)": capital,
        "Tasa semanal (%)": round(tasa * 100, 3),
        "N° Cuotas": cuotas,
        "Cuota fija (S/)": round(cuota_fija, 2),
        "Interés Total (S/)": round(interes_total, 2),
        "Degravamen Total (S/)": degrav_total,
        "Total a Pagar (S/)": round(suma_cuotas_final, 2)
    }

    return df, resumen

# --------------------------
# UTIL: EXPORTAR A XLSX
# --------------------------
def to_excel_bytes(df, resumen, filename="cronograma.xlsx"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Cronograma")
        res_df = pd.DataFrame(list(resumen.items()), columns=["Concepto", "Valor"])
        res_df.to_excel(writer, index=False, sheet_name="Resumen")
    return output.getvalue()

# --------------------------
# LOGO BASE64
# --------------------------
def _get_logo_base64():
    import base64, os
    if os.path.exists(LOGO_PATH):
        with open(LOGO_PATH, "rb") as f:
            data = f.read()
        return base64.b64encode(data).decode("utf-8")
    else:
        return ""

# --------------------------
# HEADER
# --------------------------
def header():
    st.markdown(
        f"""
        <style>
        .abecoin-header {{
            background: linear-gradient(90deg, {AZUL}, {AZUL});
            padding: 18px;
            border-radius: 8px;
            color: {TEXTO_HEADER};
            display: flex;
            align-items: center;
            gap: 20px;
        }}
        .abecoin-logo {{
            height: 70px;
            width: auto;
            border-radius: 6px;
        }}
        .abecoin-title {{
            font-size:32px;
            font-weight:700;
            margin:0;
        }}
        </style>
        <div class="abecoin-header">
            <img src="data:image/png;base64,{_get_logo_base64()}" class="abecoin-logo"/>
            <div>
                <p class="abecoin-title">ABECOIN</p>
                <div style="font-size:14px;">Simulador de préstamos — Interés Compuesto</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

# --------------------------
# INTERFAZ PRINCIPAL
# --------------------------
header()

col1, col2 = st.columns([2, 3])

with col1:
    st.markdown("### 📋 Datos del socio")
    with st.form("datos_form"):
        nombre = st.text_input("Nombre completo")
        dni = st.text_input("DNI")
        direccion = st.text_input("Dirección")
        col_a, col_b = st.columns(2)
        with col_a:
            capital = st.number_input("Monto del préstamo (S/)", min_value=10.0, step=10.0, value=200.0)
        with col_b:
            cuotas = st.number_input("N° de cuotas (semanas)", min_value=1, step=1, value=3)
        degrav_mode = st.selectbox("Cómo cobrar Degravamen?", ("prorated", "upfront"))
        submitted = st.form_submit_button("Calcular Cronograma")

    st.markdown("---")
    st.markdown("#### 🧾 Vista previa")
    if submitted and nombre and dni:
        df_preview, resumen_preview = generar_cronograma(nombre, dni, direccion, capital, int(cuotas), degrav_mode)
        st.metric("Total a pagar (S/)", f"{resumen_preview['Total a Pagar (S/)']}")
        st.write(f"Cuota fija (S/): {resumen_preview['Cuota fija (S/)']}")
        st.write(f"Interés total (S/): {resumen_preview['Interés Total (S/)']}")
        st.write(f"Degravamen total (S/): {resumen_preview['Degravamen Total (S/)']}")
    else:
        st.info("Complete el formulario y presione 'Calcular Cronograma' para ver resultados.")

with col2:
    st.markdown("### 📅 Cronograma")
    if submitted and nombre and dni:
        df, resumen = generar_cronograma(nombre, dni, direccion, capital, int(cuotas), degrav_mode)
        st.dataframe(df, use_container_width=True)

        csv = df.to_csv(index=False).encode('utf-8')
        st.download_button("⬇ Descargar CSV", data=csv, file_name=f"cronograma_{dni}.csv", mime="text/csv")

        xlsx_bytes = to_excel_bytes(df, resumen, filename=f"cronograma_{dni}.xlsx")
        st.download_button("⬇ Descargar Excel (XLSX)", data=xlsx_bytes, file_name=f"cronograma_{dni}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.markdown("#### Resumen detallado")
        for k, v in resumen.items():
            st.write(f"**{k}:** {v}")

        st.markdown("#### 📈 Evolución del saldo")
        chart_df = df[["N° Cuota", "Saldo Capital"]].set_index("N° Cuota")
        st.line_chart(chart_df)
    else:
        st.write("Aquí aparecerá el cronograma una vez ingreses los datos y presiones calcular.")

# --------------------------
# SIDEBAR
# --------------------------
st.sidebar.image(LOGO_PATH, width=120)
st.sidebar.markdown("# ABECOIN")
st.sidebar.markdown("Cooperativa de Ahorro y Crédito")
st.sidebar.markdown("---")
st.sidebar.markdown("### Contacto")
st.sidebar.write("📧 Abecooin@gmail.com")
st.sidebar.write("📞 +51 957 607 754")
st.sidebar.markdown("---")
st.sidebar.markdown("### Recomendaciones")
st.sidebar.write("- Usa montos reales para mejores resultados.")
st.sidebar.write("- Elige prorrateado si quieres cuotas estables.")



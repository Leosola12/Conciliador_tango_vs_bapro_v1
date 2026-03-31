import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Conciliador Bancario", layout="wide")

st.title("💼 Conciliador Mayor vs Extracto")

# Uploads
archivo_extracto = st.file_uploader("Subí el EXTRACTO bancario", type=["xlsx"])
archivo_mayor = st.file_uploader("Subí el MAYOR contable", type=["xlsx"])

tolerancia_dias = st.number_input("Tolerancia de días", value=2)

if archivo_extracto and archivo_mayor:

    extracto = pd.read_excel(archivo_extracto)
    mayor = pd.read_excel(archivo_mayor)

    st.write("Columnas extracto:", extracto.columns.tolist())
    st.write("Columnas mayor:", mayor.columns.tolist())

    try:
        # === NORMALIZAR NOMBRES ===
        cols_ext_lower = [c.lower().strip() for c in extracto.columns]
        cols_may_lower = [c.lower().strip() for c in mayor.columns]

        # === VALIDAR ===
        if "fecha" not in cols_ext_lower:
            st.error("❌ Extracto sin columna Fecha")
            st.stop()

        tiene_importe = "importe" in cols_ext_lower
        tiene_debito = ("débito" in cols_ext_lower) or ("debito" in cols_ext_lower)
        tiene_credito = ("crédito" in cols_ext_lower) or ("credito" in cols_ext_lower)

        if not tiene_importe and not (tiene_debito and tiene_credito):
            st.error("❌ Extracto inválido")
            st.stop()

        for c in ["fecha", "debe", "haber"]:
            if c not in cols_may_lower:
                st.error(f"❌ Falta columna {c} en mayor")
                st.stop()

        # === MAYOR ===
        mayor["Debe"] = pd.to_numeric(mayor["Debe"], errors="coerce").fillna(0)
        mayor["Haber"] = pd.to_numeric(mayor["Haber"], errors="coerce").fillna(0)
        mayor["Importe"] = mayor["Debe"] - mayor["Haber"]

        # === EXTRACTO ===
        cols = {c.lower().strip(): c for c in extracto.columns}

        tiene_importe = "importe" in cols
        tiene_debito = "débito" in cols or "debito" in cols
        tiene_credito = "crédito" in cols or "credito" in cols

        if tiene_debito and tiene_credito:
            deb_col = cols.get("débito", cols.get("debito"))
            cred_col = cols.get("crédito", cols.get("credito"))

            extracto["Importe"] = (
                pd.to_numeric(extracto[cred_col], errors="coerce").fillna(0)
                +
                pd.to_numeric(extracto[deb_col], errors="coerce").fillna(0)
            )

        elif tiene_importe:
            imp_col = cols["importe"]
            extracto["Importe"] = pd.to_numeric(
                extracto[imp_col], errors="coerce"
            ).fillna(0)

        else:
            st.error("❌ No se pudo construir Importe")
            st.stop()

        # === FECHAS ===
        extracto["Fecha"] = pd.to_datetime(extracto["Fecha"], dayfirst=True, errors="coerce")
        mayor["Fecha"] = pd.to_datetime(mayor["Fecha"], dayfirst=True, errors="coerce")

        # === MATCH ===
        def buscar_match(row):
            posibles = mayor[
                (abs(mayor["Importe"] - row["Importe"]) < 0.01) &
                (abs((mayor["Fecha"] - row["Fecha"]).dt.days) <= tolerancia_dias)
            ]
            return len(posibles) > 0

        extracto["Match"] = extracto.apply(buscar_match, axis=1)

        ok = extracto[extracto["Match"] == True]
        solo_banco = extracto[extracto["Match"] == False]
        solo_mayor = mayor[~mayor["Importe"].isin(extracto["Importe"])]

        st.success("✅ Conciliación completada")

        col1, col2, col3 = st.columns(3)
        col1.metric("Conciliados", len(ok))
        col2.metric("Solo banco", len(solo_banco))
        col3.metric("Solo mayor", len(solo_mayor))

        st.subheader("Vista previa")
        st.dataframe(ok.head(20))

        # === EXPORTAR ===
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            ok.to_excel(writer, sheet_name="Conciliados", index=False)
            solo_banco.to_excel(writer, sheet_name="Solo_Banco", index=False)
            solo_mayor.to_excel(writer, sheet_name="Solo_Mayor", index=False)

        st.download_button(
            label="📥 Descargar Excel",
            data=output.getvalue(),
            file_name="Conciliacion.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {e}")
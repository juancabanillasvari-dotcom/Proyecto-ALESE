import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Seguimiento de Unidades - ALESE", layout="wide")

st.title("🚗 Control de Estado de Unidades — Cliente ALESE")
st.markdown("Sube tu archivo Excel de operaciones para procesarlo y obtener los reportes clasificados en **Finalizado**, **En proceso** y **En cola**.")

uploaded_file = st.file_uploader("Selecciona o arrastra tu archivo Excel", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    # Filtrar únicamente los registros del cliente ALESE
    df_alese = df[df['Cliente'].astype(str).str.contains('ALESE', case=False, na=False)].copy()

    if df_alese.empty:
        st.warning("No se encontraron registros para el cliente ALESE en el archivo subido.")
    else:
        # Lógica de clasificación (PDI Inicial -> PDI Final, sin pulido básico)
        def clasificar_estado(row):
            ubicacion = str(row.get('UBICACIÓN VARI', '')).upper()
            pdi_inicial = str(row.get('PDI Inicial', '')).upper()
            pdi_final = str(row.get('PDI Final', '')).upper()

            if "FINALIZADO" in ubicacion or pdi_final == "COMPLETO":
                return "Finalizado"
            elif "PDI FINAL" in ubicacion or "ZONA DESPACHO" in ubicacion or pdi_final == "EN PROCESO":
                return "En proceso"
            else:
                return "En cola"

        df_alese['Estado General'] = df_alese.apply(clasificar_estado, axis=1)

        # Métricas visuales
        st.markdown("---")
        col1, col2, col3 = st.columns(3)
        col1.metric("🟢 Finalizados", (df_alese['Estado General'] == "Finalizado").sum())
        col2.metric("🟡 En Proceso", (df_alese['Estado General'] == "En proceso").sum())
        col3.metric("🔴 En Cola", (df_alese['Estado General'] == "En cola").sum())
        st.markdown("---")

        # Vista previa en tabla
        st.subheader("Vista Previa de Unidades Procesadas")
        st.dataframe(df_alese, use_container_width=True)

        # Generar Excel para descarga
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_alese.to_excel(writer, index=False, sheet_name='Unidades ALESE')
        
        st.download_button(
            label="📥 Descargar Reporte Excel Procesado",
            data=output.getvalue(),
            file_name="Estado_Unidades_ALESE_Procesado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

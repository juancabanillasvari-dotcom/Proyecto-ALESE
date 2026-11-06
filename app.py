import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Seguimiento de Unidades - ALESE", layout="wide")

st.title("🚗 Control de Estado de Unidades — Cliente ALESE")
st.markdown("Sube tu reporte nativo de Excel para obtener el estado clasificado, correlativo y ordenado estrictamente de más antiguo a más actual.")

uploaded_file = st.file_uploader("Sube tu archivo Excel (ej. Automate_diario.xlsx)", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    else:
        sheet_target = 'SeguimientoServicio' if 'SeguimientoServicio' in pd.ExcelFile(uploaded_file).sheet_names else 0
        df = pd.read_excel(uploaded_file, sheet_name=sheet_target)

    # Filtrar solo cliente ALESE y unidades que NO estén RETIRADAS
    df_alese = df[
        (df['Cliente'].astype(str).str.contains('ALESE', case=False, na=False)) &
        (df['Ubicacion'].astype(str).str.upper() != 'RETIRADO')
    ].copy()

    if df_alese.empty:
        st.warning("No se encontraron unidades en almacén para el cliente ALESE en el archivo subido.")
    else:
        piv = df_alese.pivot_table(
            index=['Cliente', 'Almacen', 'Vin', 'Marca', 'Modelo', 'Color Vehiculo'],
            columns='SubServicio',
            values=['Estado Vehiculo', 'Fecha Termino', 'Fecha Ingreso'],
            aggfunc='first'
        )
        piv.columns = [f'{col[0]}_{col[1]}' for col in piv.columns]
        piv = piv.reset_index()

        def evaluar_estado_unidad(row):
            pdi_final_est = str(row.get('Estado Vehiculo_PDI FINAL', '')).upper()
            pdi_inicial_est = str(row.get('Estado Vehiculo_PDI INICIAL COMPLETO', '')).upper()

            if pdi_final_est.startswith("FINALIZADO"):
                return "FINALIZADO"
            if "EN COLA" in pdi_inicial_est or pdi_inicial_est in ["NAN", "NONE", ""]:
                return "EN COLA"
            if "EN COLA" in pdi_final_est or "EN PROCESO" in pdi_final_est or "EN PROCESO" in pdi_inicial_est:
                return "EN PROCESO"
            return "EN COLA"

        piv['Estado Vehiculo'] = piv.apply(evaluar_estado_unidad, axis=1)

        # Asignar fecha cruda para ordenamiento estricto cronológico
        def asignar_fecha_raw(row):
            estado = row['Estado Vehiculo']
            if estado in ["EN COLA", "EN PROCESO"]:
                return row.get('Fecha Ingreso_PDI INICIAL COMPLETO', row.get('Fecha Ingreso_PDI FINAL', None))
            else:
                return row.get('Fecha Termino_PDI FINAL', row.get('Fecha Termino_PDI INICIAL COMPLETO', None))

        piv['Fecha_Raw'] = piv.apply(asignar_fecha_raw, axis=1)
        
        # Convertir a datetime interpretando primero el día (dayfirst=True) para orden cronológico real
        piv['Fecha_Datetime'] = pd.to_datetime(piv['Fecha_Raw'], errors='coerce', dayfirst=True)

        # Extraer exactamente los 10 primeros caracteres del archivo nativo de izquierda a derecha para mostrar
        def extraer_10_chars(row):
            val = row['Fecha_Raw']
            if pd.isna(val) or str(val).strip() in ["", "nan", "NaT", "None"]:
                return ""
            return str(val).strip()[:10]

        piv['Fecha'] = piv.apply(extraer_10_chars, axis=1)

        # Ordenamiento estrictamente de menor a mayor (más antiguo arriba, ej: 29/10 antes que 05/11)
        estado_order = {"FINALIZADO": 1, "EN PROCESO": 2, "EN COLA": 3}
        piv['Orden_Estado'] = piv['Estado Vehiculo'].map(estado_order)
        piv = piv.sort_values(by=['Orden_Estado', 'Fecha_Datetime'], ascending=[True, True])

        cols_export = ['Cliente', 'Almacen', 'Vin', 'Marca', 'Modelo', 'Color Vehiculo', 'Estado Vehiculo', 'Fecha']
        df_final = piv[cols_export]

        # Métricas principales
        st.markdown("---")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("📦 Total en Almacén", len(df_final))
        m2.metric("🟢 Finalizadas", (df_final['Estado Vehiculo'] == "FINALIZADO").sum())
        m3.metric("🟡 En Proceso", (df_final['Estado Vehiculo'] == "EN PROCESO").sum())
        m4.metric("🔴 En Cola", (df_final['Estado Vehiculo'] == "EN COLA").sum())
        st.markdown("---")

        # Preparar dataframes con numeración correlativa N°
        df_fin = df_final[df_final['Estado Vehiculo'] == "FINALIZADO"].drop(columns=['Estado Vehiculo']).copy()
        df_fin.insert(0, 'N°', range(1, len(df_fin) + 1))

        df_proc = df_final[df_final['Estado Vehiculo'] == "EN PROCESO"].drop(columns=['Estado Vehiculo']).copy()
        df_proc.insert(0, 'N°', range(1, len(df_proc) + 1))

        df_cola = df_final[df_final['Estado Vehiculo'] == "EN COLA"].drop(columns=['Estado Vehiculo']).copy()
        df_cola.insert(0, 'N°', range(1, len(df_cola) + 1))

        st.subheader("🟢 Unidades Finalizadas")
        st.dataframe(df_fin, use_container_width=True)

        st.subheader("🟡 Unidades En Proceso")
        st.dataframe(df_proc, use_container_width=True)

        st.subheader("🔴 Unidades En Cola")
        st.dataframe(df_cola, use_container_width=True)

        # Generar Excel estructurado con 3 pestañas independientes para descarga
        output = io.BytesIO()
        with pd.

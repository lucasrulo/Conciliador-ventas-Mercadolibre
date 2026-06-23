import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO

st.set_page_config(page_title="Conciliador E-commerce: MELI - MP - SAP", layout="wide")

st.title("📊 Conciliador de Operaciones E-commerce")
st.subheader("Mercado Libre | Mercado Pago | SAP")

# -----------------------------------------------------------------------------
# Configuración de Cuentas por Marca
# -----------------------------------------------------------------------------
BRAND_ACCOUNTS = {
    '1.1.040.60.000': 'CROCS',
    '1.1.040.50.000': 'KAPPA',
    '1.1.040.80.000': 'PICADILLY',
    '1.1.040.95.000': 'COLUMBIA',
    '1.1.040.92.000': 'REEBOK'
}

# -----------------------------------------------------------------------------
# Funciones de Limpieza y Procesamiento
# -----------------------------------------------------------------------------
def extraer_pack_id_sap(v):
    if pd.isna(v):
        return None
    v_str = str(v).strip()
    # Busca la secuencia de números al final de la cadena (ej. MELI_REEBOK2000013588172475)
    match = re.search(r'\d+$', v_str)
    return match.group(0) if match else v_str

def detectar_marca_sap_texto(v):
    if pd.isna(v):
        return "DESCONOCIDO"
    v_str = str(v).upper()
    for brand in ['CROCS', 'KAPPA', 'PICADILLY', 'COLUMBIA', 'REEBOK']:
        if brand in v_str:
            return brand
    return "DESCONOCIDO"

# -----------------------------------------------------------------------------
# Interfaz de Carga de Archivos
# -----------------------------------------------------------------------------
st.sidebar.header("📁 Carga de Reportes Excel")
file_meli = st.sidebar.file_uploader("1. Reporte Mercado Libre (MELI)", type=["xlsx", "xls"])
file_mp = st.sidebar.file_uploader("2. Reporte Mercado Pago (MP)", type=["xlsx", "xls"])
file_sap = st.sidebar.file_uploader("3. Reporte SAP", type=["xlsx", "xls"])

if file_meli and file_mp and file_sap:
    with st.spinner("Procesando y cruzando los datos..."):
        try:
            # 1. LEER ARCHIVOS
            df_meli_raw = pd.read_excel(file_meli)
            df_mp_raw = pd.read_excel(file_mp)
            df_sap_raw = pd.read_excel(file_sap)

            # --- PROCESAMIENTO MERCADO LIBRE ---
            df_meli = df_meli_raw.copy()
            # Asegurar nombres de columnas limpios
            df_meli.columns = [str(c).strip() for c in df_meli.columns]
            
            # Identificar columnas clave
            col_meli_id = '# de venta'
            col_meli_fecha = 'Fecha de venta'
            col_meli_monto = 'Ingresos por productos (ARS)'
            
            df_meli['Pack_ID_Clean'] = df_meli[col_meli_id].astype(str).str.strip()
            df_meli['Fecha_Clean'] = pd.to_datetime(df_meli[col_meli_fecha]).dt.date
            df_meli['Monto_MELI'] = pd.to_numeric(df_meli[col_meli_monto], errors='coerce').fillna(0)

            # --- PROCESAMIENTO MERCADO PAGO ---
            df_mp = df_mp_raw.copy()
            df_mp.columns = [str(c).strip() for c in df_mp.columns]
            
            col_mp_ref = 'Código de referencia (external_reference)'
            col_mp_fecha = 'Fecha de compra (date_created)'
            col_mp_order = 'Número de venta en Mercado Libre (order_id)'
            
            df_mp['Ext_Ref_Clean'] = df_mp[col_mp_ref].astype(str).str.strip()
            df_mp['Order_ID_Clean'] = df_mp[col_mp_order].astype(str).str.strip()
            df_mp['Fecha_Clean'] = pd.to_datetime(df_mp[col_mp_fecha]).dt.date
            
            # Buscar columna de monto (adaptable si cambia el nombre exacto de la columna de dinero neta)
            col_mp_monto = [c for c in df_mp.columns if 'monto' in c.lower() or 'importe' in c.lower() or 'neto' in c.lower() or 'total' in c.lower()]
            col_mp_monto = col_mp_monto[0] if col_mp_monto else df_mp.columns[len(df_mp.columns)-1] 
            df_mp['Monto_MP'] = pd.to_numeric(df_mp[col_mp_monto], errors='coerce').fillna(0)

            # --- PROCESAMIENTO SAP ---
            df_sap = df_sap_raw.copy()
            df_sap.columns = [str(c).strip() for c in df_sap.columns]
            
            col_sap_vtex = 'Pedido VTEX o marca'
            col_sap_cuenta = 'Cta.efectivo'
            col_sap_fecha = 'Fecha de contabilización'
            
            # Buscar columna de monto en SAP
            col_sap_monto = [c for c in df_sap.columns if 'importe' in c.lower() or 'monto' in c.lower() or 'saldo' in c.lower() or 'total' in c.lower()]
            col_sap_monto = col_sap_monto[0] if col_sap_monto else df_sap.columns[len(df_sap.columns)-1]

            df_sap['Pack_ID_Clean'] = df_sap[col_sap_vtex].apply(extraer_pack_id_sap)
            df_sap['Marca_Texto_SAP'] = df_sap[col_sap_vtex].apply(detectar_marca_sap_texto)
            df_sap['Cuenta_Clean'] = df_sap[col_sap_cuenta].astype(str).str.strip()
            df_sap['Marca_Cuenta_SAP'] = df_sap['Cuenta_Clean'].map(BRAND_ACCOUNTS).fillna("OTRA / DESCONOCIDA")
            df_sap['Fecha_Clean'] = pd.to_datetime(df_sap[col_sap_fecha]).dt.date
            df_sap['Monto_SAP'] = pd.to_numeric(df_sap[col_sap_monto], errors='coerce').fillna(0)

            # -----------------------------------------------------------------
            # LÓGICA DE CONCILIACIÓN (Cruces robustos)
            # -----------------------------------------------------------------
            
            # Como MP viene abierto por Order ID y SAP/MELI consolidan en Pack ID,
            # usaremos los datos disponibles para mapear Order ID -> Pack ID.
            # En muchos casos, si external_reference o el mapeo interno coincide, agrupamos MP.
            
            # Agrupamos MP por el external_reference (que suele agrupar o mapearse al Pack)
            # Si external_reference no coincide con Pack ID directo, creamos un consolidado por Ext_Ref_Clean
            mp_grouped = df_mp.groupby('Ext_Ref_Clean').agg({
                'Monto_MP': 'sum',
                'Fecha_Clean': 'first',
                'Order_ID_Clean': lambda x: ", ".join(x.unique())
            }).reset_index()
            
            # Cambiamos nombre para cruzar con el Pack ID maestro
            mp_grouped.rename(columns={'Ext_Ref_Clean': 'Pack_ID_Clean'}, inplace=True)

            # Agrupamos MELI por Pack ID por si hubiera líneas duplicadas
            meli_grouped = df_meli.groupby('Pack_ID_Clean').agg({
                'Monto_MELI': 'sum',
                'Fecha_Clean': 'first'
            }).reset_index()

            # Agrupamos SAP por Pack ID
            sap_grouped = df_sap.groupby('Pack_ID_Clean').agg({
                'Monto_SAP': 'sum',
                'Fecha_Clean': 'first',
                'Cuenta_Clean': 'first',
                'Marca_Texto_SAP': 'first',
                'Marca_Cuenta_SAP': 'first'
            }).reset_index()

            # UNIVERSO TOTAL DE PACK IDs
            all_packs = set(meli_grouped['Pack_ID_Clean']).union(set(mp_grouped['Pack_ID_Clean'])).union(set(sap_grouped['Pack_ID_Clean']))
            df_master = pd.DataFrame({'Pack_ID_Clean': list(all_packs)})

            # MERGE MAESTRO
            df_master = df_master.merge(meli_grouped, on='Pack_ID_Clean', how='left')
            df_master = df_master.merge(mp_grouped, on='Pack_ID_Clean', how='left', suffixes=('_MELI', '_MP'))
            df_master = df_master.merge(sap_grouped, on='Pack_ID_Clean', how='left')
            df_master.rename(columns={'Fecha_Clean': 'Fecha_SAP'}, inplace=True)

            # -----------------------------------------------------------------
            # VALIDACIONES REQUERIDAS
            # -----------------------------------------------------------------
            
            # 1. Falta PR en SAP
            df_master['Falta_En_SAP'] = df_master['Monto_SAP'].isna() | (df_master['Monto_SAP'] == 0)
            
            # 2. Diferencias de montos
            df_master['Diff_MELI_MP'] = (df_master['Monto_MELI'].fillna(0) - df_master['Monto_MP'].fillna(0)).round(2)
            df_master['Diff_MP_SAP'] = (df_master['Monto_MP'].fillna(0) - df_master['Monto_SAP'].fillna(0)).round(2)
            df_master['Tiene_Diff_Monto'] = (df_master['Diff_MELI_MP'] != 0) | (df_master['Diff_MP_SAP'] != 0)

            # 3. Mismas fechas migradas
            def verificar_fechas(row):
                fechas = []
                if not pd.isna(row['Fecha_Clean_MELI']): fechas.append(row['Fecha_Clean_MELI'])
                if not pd.isna(row['Fecha_Clean_MP']): fechas.append(row['Fecha_Clean_MP'])
                if not pd.isna(row['Fecha_SAP']): fechas.append(row['Fecha_SAP'])
                
                if len(set(fechas)) > 1:
                    return "Descalce de Fecha"
                return "OK"

            df_master['Control_Fecha'] = df_master.apply(verificar_fechas, axis=1)

            # 4. Cuenta de marca correspondiente en SAP
            def verificar_marca(row):
                if pd.isna(row['Monto_SAP']) or row['Monto_SAP'] == 0:
                    return "Sin Registro en SAP"
                if row['Marca_Texto_SAP'] == "DESCONOCIDO":
                    return "Marca no identificada en Texto SAP"
                if row['Marca_Texto_SAP'] != row['Marca_Cuenta_SAP']:
                    return f"Error: Texto dice {row['Marca_Texto_SAP']} pero imputó a cuenta de {row['Marca_Cuenta_SAP']}"
                return "OK"

            df_master['Control_Marca_SAP'] = df_master.apply(verificar_marca, axis=1)

            # 5. Pagos Duplicados
            # Detectar si un Pack ID aparece repetido en los reportes base antes de agrupar
            dup_meli = df_meli_raw[df_meli_raw.duplicated(subset=[col_meli_id], keep=False)]
            # Para MP excluimos las aperturas normales por producto buscando si repite mismo order_id con mismo item_id de forma idéntica
            dup_mp = df_mp_raw[df_mp_raw.duplicated(subset=[col_mp_order, 'Identificador de producto (item_id)'], keep=False)] if 'Identificador de producto (item_id)' in df_mp_raw.columns else pd.DataFrame()
            dup_sap = df_sap_raw[df_sap_raw.duplicated(subset=[col_sap_vtex], keep=False)]

            # -----------------------------------------------------------------
            # VISUALIZACIÓN EN STREAMLIT
            # -----------------------------------------------------------------
            
            # Indicadores Clave (KPIs)
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total Packs Analizados", len(df_master))
            c2.metric("Faltantes en SAP (PR)", df_master['Falta_En_SAP'].sum())
            c3.metric("Diferencias de Monto", df_master['Tiene_Diff_Monto'].sum())
            c4.metric("Errores de Imputación/Marca", (df_master['Control_Marca_SAP'].str.contains("Error")).sum())

            # Separación por Pestañas
            tab1, tab2, tab3, tab4, tab5 = st.tabs([
                "🔍 Todo el Universo", 
                "⚠️ Faltantes en SAP", 
                "💰 Diferencias de Montos", 
                "🏷️ Control de Marcas/Cuentas",
                "📅 Fechas y Duplicados"
            ])

            with tab1:
                st.dataframe(df_master[[
                    'Pack_ID_Clean', 'Monto_MELI', 'Monto_MP', 'Monto_SAP', 
                    'Fecha_Clean_MELI', 'Fecha_Clean_MP', 'Fecha_SAP', 
                    'Control_Fecha', 'Control_Marca_SAP'
                ]])

            with tab2:
                st.subheader("Órdenes pendientes de cargar/impactar en SAP")
                df_faltantes = df_master[df_master['Falta_En_SAP']]
                if not df_faltantes.empty:
                    st.dataframe(df_faltantes[['Pack_ID_Clean', 'Monto_MELI', 'Monto_MP', 'Fecha_Clean_MELI']])
                else:
                    st.success("¡Perfecto! No faltan registros en SAP.")

            with tab3:
                st.subheader("Descalces de dinero entre plataformas")
                df_diffs = df_master[df_master['Tiene_Diff_Monto']]
                if not df_diffs.empty:
                    st.dataframe(df_diffs[['Pack_ID_Clean', 'Monto_MELI', 'Monto_MP', 'Monto_SAP', 'Diff_MELI_MP', 'Diff_MP_SAP']])
                else:
                    st.success("¡Excelente! Todos los montos coinciden perfectamente.")

            with tab4:
                st.subheader("Validación de Cuentas Contables vs Marca del Pedido")
                df_marcas_err = df_master[df_master['Control_Marca_SAP'].str.contains("Error|no identificada", na=False)]
                if not df_marcas_err.empty:
                    st.dataframe(df_marcas_err[['Pack_ID_Clean', 'Monto_SAP', 'Cuenta_Clean', 'Marca_Texto_SAP', 'Marca_Cuenta_SAP', 'Control_Marca_SAP']])
                else:
                    st.success("¡Imputaciones correctas! Cada marca fue a su respectiva cuenta contable.")

            with tab5:
                st.subheader("Alertas de Fechas Descalzadas")
                df_fechas_err = df_master[df_master['Control_Control_Fecha'] == "Descalce de Fecha" if 'Control_Control_Fecha' in df_master else df_master['Control_Fecha'] == "Descalce de Fecha"]
                if not df_fechas_err.empty:
                    st.dataframe(df_fechas_err[['Pack_ID_Clean', 'Fecha_Clean_MELI', 'Fecha_Clean_MP', 'Fecha_SAP']])
                else:
                    st.info("No se encontraron descalces de días entre las plataformas.")
                
                st.subheader("Posibles Duplicados en Bases Origen")
                if not dup_sap.empty:
                    st.warning("Líneas duplicadas detectadas en el reporte SAP (Mismo Pack ID repetido):")
                    st.dataframe(dup_sap[[col_sap_vtex, col_sap_fecha, col_sap_monto]])
                if not dup_meli.empty:
                    st.warning("Líneas duplicadas detectadas en Mercado Libre:")
                    st.dataframe(dup_meli[[col_meli_id, col_meli_fecha, col_meli_monto]])

            # Exportar Resultados Combinados
            st.sidebar.markdown("---")
            st.sidebar.subheader("📥 Descargar Reporte Consolidado")
            
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_master.to_excel(writer, sheet_name='Conciliación General', index=False)
                if df_master['Falta_En_SAP'].sum() > 0:
                    df_master[df_master['Falta_En_SAP']].to_excel(writer, sheet_name='Faltantes SAP', index=False)
                if df_master['Tiene_Diff_Monto'].sum() > 0:
                    df_master[df_master['Tiene_Diff_Monto']].to_excel(writer, sheet_name='Diferencias Monto', index=False)
            
            st.sidebar.download_button(
                label="Descargar Excel de Diferencias",
                data=output.getvalue(),
                file_name="reporte_conciliacion_completo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Ocurrió un error al procesar las columnas. Asegúrate de que los archivos contengan las columnas mencionadas. Detalle: {e}")
else:
    st.info("👋 Por favor, carga los 3 archivos Excel en el panel izquierdo para comenzar la conciliación automática.")

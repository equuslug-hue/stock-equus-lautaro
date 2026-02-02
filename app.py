import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime
import time
import io

st.set_page_config(page_title="Stock Lautaro App", layout="wide")

# ==========================================
# 1. GESTIÓN DE SESIÓN Y LOGIN
# ==========================================
if 'admin_logged_in' not in st.session_state:
    st.session_state.admin_logged_in = False

def check_admin_password():
    try:
        if st.session_state.password_input == st.secrets["admin"]["password"]:
            st.session_state.admin_logged_in = True
            st.session_state.password_input = "" 
        else:
            st.error("❌ Contraseña incorrecta")
    except Exception:
        st.error("⚠️ Error: No has configurado [admin] password en los Secrets.")

def logout():
    st.session_state.admin_logged_in = False

# ==========================================
# 2. FUNCIONES DE LECTURA (CEREBRO UNIVERSAL MEJORADO)
# ==========================================

def clean_columns(df):
    """
    Limpia nombres de columnas:
    - Quita espacios
    - Elimina BOM (marcas invisibles de Excel)
    - Unifica nombres
    """
    # Limpieza agresiva de caracteres invisibles
    df.columns = [str(c).strip().replace('\ufeff', '') for c in df.columns]
    
    rename_map = {
        'Codigo': 'Código',
        'Descripcion': 'Descripción',
        'Stock 1ra un': 'Stock 1ra un.',
        'Stock': 'Stock 1ra un.',
        'Precio': 'Precio'
    }
    df.rename(columns=rename_map, inplace=True)
    return df

def try_read_csv_strategies(file_obj, strategies):
    for enc in strategies['encodings']:
        for sep in strategies['separators']:
            try:
                file_obj.seek(0)
                # Buscamos header en las primeras 20 líneas
                header_idx = 0
                found = False
                for i in range(20):
                    try:
                        line = file_obj.readline().decode(enc).lower()
                        # Buscamos variantes comunes
                        if 'código' in line or 'codigo' in line:
                            header_idx = i
                            found = True
                            break
                    except:
                        continue # Si falla decodificar la línea, seguimos
                
                if not found: header_idx = 0 

                file_obj.seek(0)
                # engine='python' es más robusto para archivos raros
                df = pd.read_csv(file_obj, header=header_idx, encoding=enc, sep=sep, engine='python')
                df = clean_columns(df)
                
                if 'Código' in df.columns:
                    return df
            except:
                continue
    return None

def load_universal_file(uploaded_file):
    filename = uploaded_file.name.lower()
    df = None
    
    # ESTRATEGIA A: Excel Nativo (.xlsx / .xls)
    # Intentamos esto primero solo si la extensión coincide, pero si falla
    # asumimos que podría ser un "Falso Excel" (HTML o CSV renombrado)
    if filename.endswith('.xlsx') or filename.endswith('.xls'):
        try:
            # Leemos sin header para buscar la fila correcta
            df_raw = pd.read_excel(uploaded_file, header=None)
            header_idx = -1
            for i, row in df_raw.head(20).iterrows():
                row_str = row.astype(str).str.lower().tolist()
                if any('código' in s or 'codigo' in s for s in row_str):
                    header_idx = i
                    break
            
            uploaded_file.seek(0)
            if header_idx != -1:
                df = pd.read_excel(uploaded_file, header=header_idx)
            else:
                df = pd.read_excel(uploaded_file)
            
            df = clean_columns(df)
            if 'Código' in df.columns:
                return df
        except Exception:
            pass # Falló lectura excel real, pasamos a CSV

    # ESTRATEGIA B: Texto / CSV (Incluye utf-8-sig para BOM)
    # Agregamos 'utf-8-sig' que es crucial para Excels guardados como CSV
    encodings = ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
    separators = [',', ';', '\t']
    
    df = try_read_csv_strategies(uploaded_file, {'encodings': encodings, 'separators': separators})
    return df

def parse_product_code(code_str):
    try:
        if pd.isna(code_str): return "S/D", "S/D", "S/D"
        parts = str(code_str).strip().split('.')
        if len(parts) < 2: return code_str, "N/A", "N/A"
        modelo = parts[0]
        suffix = parts[1]
        if len(suffix) >= 6:
            color = suffix[:3]
            talle = suffix[-3:]
        else:
            color = suffix
            talle = "N/A"
        return modelo, color, talle
    except:
        return str(code_str), "Err", "Err"

def color_rows(val):
    if '✅ OK' in str(val): return 'background-color: #d4edda; color: #155724'
    if '🔴 Falta' in str(val): return 'background-color: #f8d7da; color: #721c24'
    if '🟣 Sobra' in str(val): return 'background-color: #e2d9f3; color: #4a148c'
    return ''

# ==========================================
# 3. CONEXIÓN A GOOGLE SHEETS
# ==========================================
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data():
    try:
        df = conn.read(worksheet="Hoja1", ttl=0)
        expected_cols = ['Nombre_Lote', 'Estado_Stock', 'Código', 'Color', 'Stock 1ra un.', 'Descripción', 'Modelo_Ref']
        for col in expected_cols:
            if col not in df.columns:
                df[col] = None 
        return df
    except:
        return pd.DataFrame(columns=['Nombre_Lote', 'Estado_Stock', 'Código'])

df_master = load_data()

# ==========================================
# 4. INTERFAZ
# ==========================================
st.sidebar.title("Menú Principal")

opciones_menu = ["📂 Continuar / Editar Stock", "➕ Iniciar NUEVO Stock"]

if st.session_state.admin_logged_in:
    opciones_menu.append("🛑 Zona Admin (Borrar)")
    st.sidebar.success("🔓 Modo Admin: ON")
    if st.sidebar.button("Salir de Admin"):
        logout()
        st.rerun()
else:
    with st.sidebar.expander("🔐 Acceso Admin"):
        st.text_input("Contraseña:", type="password", key="password_input", on_change=check_admin_password)

modo = st.sidebar.radio("Ir a:", opciones_menu)

st.title("👕 Sistema de Control de Stock")

# ---------------------------------------------------------
# ZONA A: INICIAR NUEVO STOCK
# ---------------------------------------------------------
if modo == "➕ Iniciar NUEVO Stock":
    st.header("Cargar Nuevo Lote")
    col1, col2 = st.columns(2)
    with col1:
        nuevo_lote = st.text_input("Nombre del Lote / Sector", placeholder="Ej: Estanteria_A")
    with col2:
        archivo = st.file_uploader("Subir archivo", type=['xlsx', 'xls', 'csv'])

    if st.button("🚀 Procesar y Crear", type="primary") and archivo and nuevo_lote:
        nombre_limpio = nuevo_lote.strip()
        lotes_existentes = [str(x).strip() for x in df_master['Nombre_Lote'].unique()]
        
        if nombre_limpio in lotes_existentes:
            st.error(f"⚠️ Ya existe '{nombre_limpio}'. Usa otro nombre.")
        else:
            # Usamos la lectura universal mejorada
            df_new = load_universal_file(archivo)
            
            if df_new is None or 'Código' not in df_new.columns:
                st.error("❌ No se pudo leer el archivo. Verifica que tenga la columna 'Código'.")
            else:
                try:
                    parsed = df_new['Código'].apply(parse_product_code)
                    df_new['Modelo_Ref'] = [x[0] for x in parsed]
                    df_new['Color_Code'] = [x[1] for x in parsed]
                    df_new['Talle_Code'] = [x[2] for x in parsed]
                    
                    # Rellenar columnas faltantes para evitar errores
                    if 'Color' not in df_new.columns: df_new['Color'] = ""
                    if 'Stock 1ra un.' not in df_new.columns: df_new['Stock 1ra un.'] = 0
                    if 'Descripción' not in df_new.columns: df_new['Descripción'] = "S/D"
                    
                    df_new['Nombre_Lote'] = nombre_limpio
                    df_new['Estado_Stock'] = "Pendiente"
                    df_new['Cant_Diferencia'] = 0
                    df_new['Fecha_Inicio'] = datetime.now().strftime("%Y-%m-%d %H:%M")

                    df_new = df_new.dropna(axis=1, how='all')
                    updated_master = pd.concat([df_master, df_new], ignore_index=True)
                    conn.update(worksheet="Hoja1", data=updated_master)
                    
                    st.success(f"✅ Lote '{nombre_limpio}' creado.")
                    time.sleep(1.5)
                    st.rerun()
                except Exception as e:
                    st.error(f"Error procesando datos: {e}")

# ---------------------------------------------------------
# ZONA B: CONTINUAR / EDITAR (CON FILTROS EN CASCADA)
# ---------------------------------------------------------
elif modo == "📂 Continuar / Editar Stock":
    lotes = [x for x in df_master['Nombre_Lote'].unique() if str(x) != 'nan']
    
    if not lotes:
        st.info("👋 No hay stocks activos.")
    else:
        col_sel, col_btn = st.columns([3, 1])
        with col_sel:
            seleccion = st.selectbox("Selecciona el Lote:", lotes)
        with col_btn:
            st.write("") 
            st.write("") 
            if st.button("🔄 Refrescar"):
                st.cache_data.clear()
                st.rerun()
        
        # 1. Obtener Dataframe
        df_lote = df_master[df_master['Nombre_Lote'] == seleccion].copy()
        
        # 2. FILTROS EN CASCADA (Descripcion -> Modelo)
        st.write("---")
        with st.expander("🔍 Filtros de Visualización", expanded=True):
            f_col1, f_col2 = st.columns(2)
            
            with f_col1:
                # Opciones de Descripción (Siempre todas las disponibles en el lote)
                opts_desc = sorted(df_lote['Descripción'].dropna().astype(str).unique())
                sel_desc = st.multiselect("1. Filtrar por Descripción:", opts_desc, placeholder="Todos")
            
            with f_col2:
                # Opciones de Modelo (Dependen de lo que elijas en Descripción)
                if sel_desc:
                    # Si hay descripciones seleccionadas, filtramos los modelos compatibles
                    df_filtered_temp = df_lote[df_lote['Descripción'].astype(str).isin(sel_desc)]
                    opts_mod = sorted(df_filtered_temp['Modelo_Ref'].dropna().astype(str).unique())
                else:
                    # Si no, mostramos todos los modelos del lote
                    opts_mod = sorted(df_lote['Modelo_Ref'].dropna().astype(str).unique()) if 'Modelo_Ref' in df_lote.columns else []
                
                sel_mod = st.multiselect("2. Filtrar por Modelo:", opts_mod, placeholder="Todos (Filtrado por Desc.)")

        # 3. Aplicar Filtros a la Vista
        df_view = df_lote.copy()
        
        if sel_desc:
            df_view = df_view[df_view['Descripción'].astype(str).isin(sel_desc)]
            
        if sel_mod:
            df_view = df_view[df_view['Modelo_Ref'].astype(str).isin(sel_mod)]
            
        # Ordenar
        if 'Modelo_Ref' in df_view.columns:
            df_view = df_view.sort_values(by=['Modelo_Ref', 'Color_Code', 'Talle_Code'])

        st.caption(f"Mostrando {len(df_view)} items.")

        # 4. Editor
        cols_cfg = {
            "Código": st.column_config.TextColumn("Cód.", disabled=True),
            "Modelo_Ref": st.column_config.TextColumn("Modelo", disabled=True),
            "Color_Code": st.column_config.TextColumn("C.Code", disabled=True),
            "Talle_Code": st.column_config.TextColumn("Talle", disabled=True),
            "Color": st.column_config.TextColumn("Color Real", disabled=True),
            "Stock 1ra un.": st.column_config.NumberColumn("Stock Sist.", disabled=True),
            "Descripción": st.column_config.TextColumn("Desc.", disabled=True),
            "Estado_Stock": st.column_config.SelectboxColumn("Estado", options=["Pendiente", "✅ OK", "🔴 Falta", "🟣 Sobra"], required=True),
            "Cant_Diferencia": st.column_config.NumberColumn("Dif (+/-)", min_value=0, step=1)
        }
        
        cols_show = [
            'Modelo_Ref', 'Color_Code', 'Talle_Code', 'Color', 
            'Descripción', 'Stock 1ra un.', 'Estado_Stock', 'Cant_Diferencia'
        ]
        cols_final = [c for c in cols_show if c in df_view.columns]

        edited_view = st.data_editor(
            df_view[cols_final],
            column_config=cols_cfg,
            use_container_width=True,
            height=500,
            hide_index=True,
            key=f"editor_{seleccion}"
        )
        
        # 5. Guardado
        if st.button("💾 Guardar Cambios en la Nube", type="primary"):
            df_lote.update(edited_view)
            
            df_clean_master = df_master[df_master['Nombre_Lote'] != seleccion]
            df_final_upload = pd.concat([df_clean_master, df_lote], ignore_index=True)
            
            conn.update(worksheet="Hoja1", data=df_final_upload)
            st.success("✅ Guardado sincronizado.")
            time.sleep(1.5)
            st.rerun()

        # Resumen (Muestra diferencias de TODO el lote, no solo lo filtrado)
        st.divider()
        diferencias = df_lote[df_lote['Estado_Stock'].isin(['🔴 Falta', '🟣 Sobra'])]
        if not diferencias.empty:
            st.write("### 🚨 Diferencias en el Lote Completo")
            st.dataframe(diferencias[cols_final].style.map(color_rows, subset=['Estado_Stock']), use_container_width=True, hide_index=True)
        else:
            st.info("Todo OK o Pendiente en el reporte general.")

# ---------------------------------------------------------
# ZONA C: ADMIN
# ---------------------------------------------------------
elif modo == "🛑 Zona Admin (Borrar)":
    st.header("⚠️ Zona de Peligro")
    lotes = [x for x in df_master['Nombre_Lote'].unique() if str(x) != 'nan']
    if not lotes:
        st.write("Nada para borrar.")
    else:
        lote_borrar = st.selectbox("Eliminar lote:", lotes)
        if st.button(f"🔥 ELIMINAR '{lote_borrar}'"):
            df_nuevo_master = df_master[df_master['Nombre_Lote'] != lote_borrar]
            conn.update(worksheet="Hoja1", data=df_nuevo_master)
            st.success("Eliminado.")
            time.sleep(1)
            st.rerun()

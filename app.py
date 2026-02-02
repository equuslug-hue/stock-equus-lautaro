import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime
import time

st.set_page_config(page_title="Stock Lautaro App", layout="wide")

# ==========================================
# 1. GESTIÓN DE SESIÓN Y LOGIN
# ==========================================
if 'admin_logged_in' not in st.session_state:
    st.session_state.admin_logged_in = False

def check_admin_password():
    """Verifica la contraseña contra los secrets de Streamlit"""
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
# 2. FUNCIONES DE LECTURA DE ARCHIVOS (LA CLAVE)
# ==========================================

def clean_columns(df):
    """Limpia nombres de columnas y unifica 'Codigo' a 'Código'"""
    df.columns = [str(c).strip() for c in df.columns]
    if 'Codigo' in df.columns:
        df.rename(columns={'Codigo': 'Código'}, inplace=True)
    return df

def load_universal_file(uploaded_file):
    """
    Intenta leer Excel (.xlsx) o CSV (.csv) buscando la cabecera automáticamente.
    """
    filename = uploaded_file.name.lower()
    
    # --- ESTRATEGIA 1: ES UN EXCEL (.xlsx / .xls) ---
    if filename.endswith('.xlsx') or filename.endswith('.xls'):
        try:
            # Leemos sin header primero para buscar dónde empieza
            df_raw = pd.read_excel(uploaded_file, header=None)
            
            # Buscamos en las primeras 10 filas dónde está la palabra "Código"
            header_idx = -1
            for i, row in df_raw.head(15).iterrows():
                # Convertimos la fila a string y buscamos la palabra clave
                row_str = row.astype(str).str.lower().tolist()
                if any('código' in s or 'codigo' in s for s in row_str):
                    header_idx = i
                    break
            
            if header_idx != -1:
                # Recargamos usando esa fila como cabecera
                uploaded_file.seek(0)
                df = pd.read_excel(uploaded_file, header=header_idx)
                df = clean_columns(df)
                return df
            else:
                # Si no encuentra header, intenta lectura estándar
                uploaded_file.seek(0)
                df = pd.read_excel(uploaded_file)
                df = clean_columns(df)
                return df
        except Exception as e:
            st.error(f"Error leyendo Excel: {e}")
            return None

    # --- ESTRATEGIA 2: ES UN CSV (Texto plano) ---
    else:
        encodings = ['utf-8', 'latin-1', 'cp1252']
        separators = [',', ';', '\t']
        
        for enc in encodings:
            try:
                # Búsqueda manual de header en texto
                uploaded_file.seek(0)
                header_idx = 0
                found = False
                for i in range(15):
                    line = uploaded_file.readline().decode(enc).lower()
                    if 'código' in line or 'codigo' in line:
                        header_idx = i
                        found = True
                        break
                
                if not found: header_idx = 0

                # Probar separadores
                for sep in separators:
                    uploaded_file.seek(0)
                    try:
                        df = pd.read_csv(uploaded_file, header=header_idx, encoding=enc, sep=sep)
                        df = clean_columns(df)
                        if 'Código' in df.columns:
                            return df
                    except:
                        continue
            except:
                continue
                
    return None

def parse_product_code(code_str):
    """Separa Modelo, Color y Talle"""
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
        if 'Nombre_Lote' not in df.columns:
            return pd.DataFrame(columns=['Nombre_Lote', 'Estado_Stock', 'Código'])
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
    st.header("Cargar Nuevo Lote (Excel o CSV)")
    
    col1, col2 = st.columns(2)
    with col1:
        nuevo_lote = st.text_input("Nombre del Lote / Sector", placeholder="Ej: Deposito_Estanteria_1")
    with col2:
        archivo = st.file_uploader("Subir archivo (.xlsx o .csv)", type=['xlsx', 'xls', 'csv'])

    if st.button("🚀 Procesar y Crear", type="primary") and archivo and nuevo_lote:
        # Limpieza de nombre y validación estricta
        nombre_limpio = nuevo_lote.strip()
        lotes_existentes = [str(x).strip() for x in df_master['Nombre_Lote'].unique()]
        
        if nombre_limpio in lotes_existentes:
            st.error(f"⚠️ Ya existe un lote llamado '{nombre_limpio}'. Usa otro nombre.")
        else:
            # --- USAMOS LA NUEVA FUNCIÓN UNIVERSAL ---
            df_new = load_universal_file(archivo)
            
            if df_new is None or 'Código' not in df_new.columns:
                st.error("❌ No se pudo leer el archivo o no se encontró la columna 'Código'. Revisa el Excel.")
            else:
                try:
                    # Procesamiento
                    parsed = df_new['Código'].apply(parse_product_code)
                    df_new['Modelo_Ref'] = [x[0] for x in parsed]
                    df_new['Color_Code'] = [x[1] for x in parsed]
                    df_new['Talle_Code'] = [x[2] for x in parsed]
                    
                    df_new['Nombre_Lote'] = nombre_limpio
                    df_new['Estado_Stock'] = "Pendiente"
                    df_new['Cant_Diferencia'] = 0
                    df_new['Fecha_Inicio'] = datetime.now().strftime("%Y-%m-%d %H:%M")

                    # Limpieza y Guardado
                    df_new = df_new.dropna(axis=1, how='all')
                    updated_master = pd.concat([df_master, df_new], ignore_index=True)
                    conn.update(worksheet="Hoja1", data=updated_master)
                    
                    st.success(f"✅ Lote '{nombre_limpio}' creado correctamente.")
                    st.balloons()
                    time.sleep(2)
                    st.rerun()
                except Exception as e:
                    st.error(f"Error procesando datos: {e}")

# ---------------------------------------------------------
# ZONA B: CONTINUAR / EDITAR
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
        
        df_lote = df_master[df_master['Nombre_Lote'] == seleccion].copy()
        
        if 'Modelo_Ref' in df_lote.columns:
            df_lote = df_lote.sort_values(by=['Modelo_Ref', 'Color_Code', 'Talle_Code'])

        total = len(df_lote)
        pendientes = len(df_lote[df_lote['Estado_Stock'] == 'Pendiente'])
        progreso = int(((total - pendientes) / total) * 100) if total > 0 else 0
        st.progress(progreso)
        st.caption(f"Progreso: {progreso}%")

        cols_cfg = {
            "Código": st.column_config.TextColumn("Código", disabled=True),
            "Descripción": st.column_config.TextColumn("Desc.", disabled=True),
            "Estado_Stock": st.column_config.SelectboxColumn("Estado", options=["Pendiente", "✅ OK", "🔴 Falta", "🟣 Sobra"], required=True),
            "Cant_Diferencia": st.column_config.NumberColumn("Cant. Dif.", min_value=0, step=1)
        }
        
        cols_show = ['Modelo_Ref', 'Color_Code', 'Talle_Code', 'Descripción', 'Estado_Stock', 'Cant_Diferencia']
        cols_final = [c for c in cols_show if c in df_lote.columns]

        edited_lote = st.data_editor(
            df_lote[cols_final],
            column_config=cols_cfg,
            use_container_width=True,
            height=500,
            hide_index=True,
            key=f"editor_{seleccion}"
        )
        
        if st.button("💾 Guardar Cambios en la Nube", type="primary"):
            df_lote.update(edited_lote)
            df_clean_master = df_master[df_master['Nombre_Lote'] != seleccion]
            df_final_upload = pd.concat([df_clean_master, df_lote], ignore_index=True)
            conn.update(worksheet="Hoja1", data=df_final_upload)
            st.success("✅ Guardado.")
            time.sleep(1)
            st.rerun()

        st.divider()
        diferencias = df_lote[df_lote['Estado_Stock'].isin(['🔴 Falta', '🟣 Sobra'])]
        if not diferencias.empty:
            st.write("### 🚨 Diferencias")
            st.dataframe(diferencias[cols_final].style.map(color_rows, subset=['Estado_Stock']), use_container_width=True, hide_index=True)
        else:
            st.info("Todo OK o Pendiente.")

# ---------------------------------------------------------
# ZONA C: ADMIN - BORRADO
# ---------------------------------------------------------
elif modo == "🛑 Zona Admin (Borrar)":
    st.header("⚠️ Zona de Peligro: Eliminar Stocks")
    lotes = [x for x in df_master['Nombre_Lote'].unique() if str(x) != 'nan']
    
    if not lotes:
        st.write("No hay stocks para borrar.")
    else:
        lote_borrar = st.selectbox("Eliminar lote:", lotes)
        st.warning(f"Se borrará '{lote_borrar}' para siempre.")
        if st.button(f"🔥 ELIMINAR '{lote_borrar}'"):
            df_nuevo_master = df_master[df_master['Nombre_Lote'] != lote_borrar]
            conn.update(worksheet="Hoja1", data=df_nuevo_master)
            st.success("Eliminado.")
            time.sleep(1)
            st.rerun()

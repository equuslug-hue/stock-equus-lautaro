import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime
import time

# Configuración básica
st.set_page_config(page_title="Stock Equus", layout="wide")

# CSS para optimización móvil
st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    /* Botones más grandes */
    .stButton > button {
        height: 3.5em; 
        font-weight: bold;
        font-size: 16px;
    }
    /* Inputs más grandes */
    .stTextInput > div > div > input {
        font-size: 16px;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 1. GESTIÓN DE SESIÓN Y MEMORIA
# ==========================================
if 'admin_logged_in' not in st.session_state: st.session_state.admin_logged_in = False
if 'cambios_temporales' not in st.session_state: st.session_state.cambios_temporales = {} # Memoria de cambios

def check_admin_password():
    try:
        if st.session_state.password_input == st.secrets["admin"]["password"]:
            st.session_state.admin_logged_in = True
            st.session_state.password_input = "" 
        else:
            time.sleep(1)
            st.error("❌ Incorrecto")
    except: st.error("⚠️ Configura secrets.")

def logout(): st.session_state.admin_logged_in = False

# ==========================================
# 2. FUNCIONES DE LECTURA
# ==========================================
def clean_columns(df):
    df.columns = [str(c).strip().replace('\ufeff', '') for c in df.columns]
    rename_map = {}
    for col in df.columns:
        c = col.lower()
        if 'codigo' in c or 'código' in c: rename_map[col] = 'Código'
        elif 'descripción' in c or 'descripcion' in c: rename_map[col] = 'Descripción'
        elif 'stock' in c: rename_map[col] = 'Stock 1ra un.'
    df.rename(columns=rename_map, inplace=True)
    return df

def load_universal_file(uploaded_file):
    filename = uploaded_file.name.lower()
    df = None
    try:
        if filename.endswith(('.xls', '.xlsx')):
            try: df = pd.read_excel(uploaded_file)
            except: pass
        
        if df is None: 
            for enc in ['utf-8-sig', 'latin-1', 'cp1252']:
                for sep in [',', ';', '\t']:
                    try:
                        uploaded_file.seek(0)
                        temp = pd.read_csv(uploaded_file, encoding=enc, sep=sep, engine='python')
                        if len(temp.columns)>1: 
                            df = temp; break
                    except: continue
                if df is not None: break
        
        if df is not None:
            df = clean_columns(df)
            return df
    except: return None
    return df

def parse_product_code(code_str):
    try:
        if pd.isna(code_str): return "S/D", "S/D", "S/D"
        p = str(code_str).strip().split('.')
        if len(p)<2: return code_str, "N/A", "N/A"
        return p[0], p[1][:3] if len(p[1])>=3 else p[1], p[1][-3:] if len(p[1])>=6 else "N/A"
    except: return str(code_str), "Err", "Err"

def sanitize_dataframe(df):
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].apply(lambda x: f"'{x}" if str(x).startswith(('=','+','-','@')) else x)
    return df

# ==========================================
# 3. CONEXIÓN Y DATOS
# ==========================================
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data():
    try:
        df = conn.read(worksheet="Hoja1", ttl=0)
        req = ['Nombre_Lote', 'Estado_Stock', 'Código', 'Color', 'Stock 1ra un.', 'Descripción', 'Modelo_Ref', 'Cant_Diferencia']
        for c in req: 
            if c not in df.columns: df[c] = None
        
        # --- APLICAR CAMBIOS TEMPORALES DE LA MEMORIA AL DATAFRAME ---
        # Esto asegura que si recargas, ves lo que editaste
        if not df.empty and st.session_state.cambios_temporales:
            for idx, cambios in st.session_state.cambios_temporales.items():
                if idx in df.index:
                    for k, v in cambios.items():
                        df.at[idx, k] = v
        return df
    except: return pd.DataFrame(columns=['Nombre_Lote'])

# Carga inicial
df_master = load_data()

# ==========================================
# 4. INTERFAZ
# ==========================================
st.sidebar.title("Menú")

modo_visual = st.sidebar.radio("Vista:", ["📱 MODO CELULAR", "💻 MODO PC"], index=0)
st.sidebar.divider()

if st.session_state.admin_logged_in:
    if st.sidebar.button("Salir Admin"): logout(); st.rerun()
    accion = st.sidebar.radio("Acción:", ["Trabajar Stock", "Nuevo Stock", "Borrar Stock"])
else:
    with st.sidebar.expander("Admin"):
        st.text_input("Clave", type="password", key="password_input", on_change=check_admin_password)
    accion = st.sidebar.radio("Acción:", ["Trabajar Stock", "Nuevo Stock"])

st.title("👕 Control de Stock")

# ---------------------------------------------------------
# A. NUEVO STOCK
# ---------------------------------------------------------
if accion == "Nuevo Stock":
    st.header("Crear Lote")
    lote = st.text_input("Nombre Sector")
    file = st.file_uploader("Archivo")
    if st.button("Procesar", type="primary", use_container_width=True) and file and lote:
        if lote.strip() in [str(x) for x in df_master['Nombre_Lote'].unique()]:
            st.error("Nombre repetido.")
        else:
            df = load_universal_file(file)
            if df is not None and 'Código' in clean_columns(df).columns:
                df = sanitize_dataframe(df)
                parsed = df['Código'].apply(parse_product_code)
                df['Modelo_Ref'] = [x[0] for x in parsed]
                df['Color_Code'] = [x[1] for x in parsed]
                df['Talle_Code'] = [x[2] for x in parsed]
                df['Nombre_Lote'] = lote.strip()
                df['Estado_Stock'] = "Pendiente"
                df['Cant_Diferencia'] = 0
                for c in ['Color','Descripción']: 
                    if c not in df.columns: df[c] = "-"
                if 'Stock 1ra un.' not in df.columns: df['Stock 1ra un.'] = 0
                
                df = df.dropna(axis=1, how='all')
                master = pd.concat([df_master, df], ignore_index=True)
                conn.update(worksheet="Hoja1", data=master)
                st.success("Creado!"); time.sleep(1); st.rerun()
            else: st.error("Error archivo.")

# ---------------------------------------------------------
# B. TRABAJAR STOCK
# ---------------------------------------------------------
elif accion == "Trabajar Stock":
    lotes = [x for x in df_master['Nombre_Lote'].unique() if str(x) != 'nan']
    if not lotes:
        st.info("Sin stocks.")
    else:
        seleccion = st.selectbox("Lote:", lotes)
        
        # Filtro de datos
        df_lote = df_master[df_master['Nombre_Lote'] == seleccion].copy()
        
        # ==================== MODO CELULAR ====================
        if modo_visual == "📱 MODO CELULAR":
            st.divider()
            
            # --- FILTROS CELULAR ---
            with st.expander("🔍 FILTROS (Toca para abrir)", expanded=False):
                f_cod = st.text_input("Buscar por Código:", placeholder="Ej: 0400...").strip()
                
                # Filtros inteligentes
                all_desc = sorted(df_lote['Descripción'].dropna().astype(str).unique())
                f_desc = st.selectbox("Filtrar Item:", ["Todos"] + all_desc)
                
                if f_desc != "Todos":
                    sub_df = df_lote[df_lote['Descripción'] == f_desc]
                    all_col = sorted(sub_df['Color'].dropna().astype(str).unique())
                else:
                    all_col = sorted(df_lote['Color'].dropna().astype(str).unique())
                
                f_col = st.selectbox("Filtrar Color:", ["Todos"] + all_col)

            # Aplicar filtros
            df_view = df_lote.copy()
            if f_cod: 
                df_view = df_view[df_view['Código'].astype(str).str.contains(f_cod, case=False, na=False)]
            if f_desc != "Todos": 
                df_view = df_view[df_view['Descripción'] == f_desc]
            if f_col != "Todos": 
                df_view = df_view[df_view['Color'] == f_col]
            
            # Ordenar
            if 'Modelo_Ref' in df_view.columns:
                df_view = df_view.sort_values(by=['Modelo_Ref', 'Color_Code', 'Talle_Code'])
            
            # --- FICHA DE PRODUCTO ---
            if len(df_view) > 0:
                # Navegación
                if 'idx_cel' not in st.session_state: st.session_state.idx_cel = 0
                if st.session_state.idx_cel >= len(df_view): st.session_state.idx_cel = 0
                
                row = df_view.iloc[st.session_state.idx_cel]
                real_idx = row.name # ID real para guardar
                
                # Estado actual (chequear memoria temporal primero)
                mem_data = st.session_state.cambios_temporales.get(real_idx, {})
                curr_est = mem_data.get('Estado_Stock', row['Estado_Stock'])
                curr_dif = float(mem_data.get('Cant_Diferencia', row['Cant_Diferencia']))
                
                # Diseño de Tarjeta
                with st.container(border=True):
                    st.caption(f"Item {st.session_state.idx_cel + 1} de {len(df_view)}")
                    st.markdown(f"#### {row['Descripción']}")
                    
                    c1, c2 = st.columns(2)
                    c1.metric("Talle", str(row.get('Talle_Code','-')))
                    c2.metric("Color", str(row.get('Color','-')))
                    
                    st.markdown(f"**Stock Sistema:** `{int(row.get('Stock 1ra un.',0))}`")
                    st.divider()
                    
                    # CONTROLES DE ESTADO (Botones grandes)
                    st.write("Estado:")
                    col_ok, col_falta, col_sobra = st.columns(3)
                    
                    # Funciones para actualizar memoria
                    def set_status(val):
                        if real_idx not in st.session_state.cambios_temporales:
                            st.session_state.cambios_temporales[real_idx] = {}
                        st.session_state.cambios_temporales[real_idx]['Estado_Stock'] = val
                        # Si pone OK, reseteamos diferencia
                        if val == "✅ OK":
                            st.session_state.cambios_temporales[real_idx]['Cant_Diferencia'] = 0
                    
                    # Renderizar botones con estilo según selección
                    if col_ok.button("✅ OK", use_container_width=True, type="primary" if curr_est=="✅ OK" else "secondary"):
                        set_status("✅ OK"); st.rerun()
                    
                    if col_falta.button("🔴 Falta", use_container_width=True, type="primary" if curr_est=="🔴 Falta" else "secondary"):
                        set_status("🔴 Falta"); st.rerun()
                        
                    if col_sobra.button("🟣 Sobra", use_container_width=True, type="primary" if curr_est=="🟣 Sobra" else "secondary"):
                        set_status("🟣 Sobra"); st.rerun()

                    # CONTROLES DIFERENCIA
                    if curr_est in ["🔴 Falta", "🟣 Sobra"]:
                        st.write("Cantidad Diferencia:")
                        cd1, cd2, cd3 = st.columns([1,2,1])
                        
                        if cd1.button("➖", key="dm", use_container_width=True):
                            st.session_state.cambios_temporales.setdefault(real_idx, {})['Cant_Diferencia'] = curr_dif - 1
                            st.rerun()
                            
                        cd2.markdown(f"<h2 style='text-align:center'>{int(curr_dif)}</h2>", unsafe_allow_html=True)
                        
                        if cd3.button("➕", key="dp", use_container_width=True):
                            st.session_state.cambios_temporales.setdefault(real_idx, {})['Cant_Diferencia'] = curr_dif + 1
                            st.rerun()

                # BOTONES DE NAVEGACIÓN
                cn1, cn2 = st.columns(2)
                if cn1.button("⬅️ Anterior", use_container_width=True):
                    st.session_state.idx_cel = max(0, st.session_state.idx_cel - 1)
                    st.rerun()
                
                # El botón siguiente funciona como "Confirmar y Avanzar"
                if cn2.button("Siguiente ➡️", use_container_width=True):
                    st.session_state.idx_cel = min(len(df_view) - 1, st.session_state.idx_cel + 1)
                    st.rerun()

            else:
                st.warning("No hay productos con esos filtros.")

            st.write("")
            st.divider()
            
            # BOTÓN DE GUARDADO FINAL (Aplica la memoria temporal a Google Sheets)
            if st.button("💾 GUARDAR TODO EN LA NUBE", type="primary", use_container_width=True):
                if st.session_state.cambios_temporales:
                    # Aplicar cambios al master
                    for idx, changes in st.session_state.cambios_temporales.items():
                        if idx in df_master.index:
                            for k, v in changes.items():
                                df_master.at[idx, k] = v
                    
                    # Subir
                    conn.update(worksheet="Hoja1", data=df_master)
                    st.session_state.cambios_temporales = {} # Limpiar memoria tras guardar
                    st.success("¡Datos guardados en Google Sheets!")
                    time.sleep(2)
                    st.rerun()
                else:
                    st.info("No hay cambios nuevos para guardar.")

        # ==================== MODO PC ====================
        else:
            # (Código PC anterior...)
            if st.button("🔄 Recargar"): st.cache_data.clear(); st.rerun()
            
            editor = st.data_editor(
                df_lote[['Talle_Code', 'Color', 'Descripción', 'Stock 1ra un.', 'Estado_Stock', 'Cant_Diferencia']],
                column_config={
                    "Estado_Stock": st.column_config.SelectboxColumn("Estado", options=["Pendiente", "✅ OK", "🔴 Falta", "🟣 Sobra"], required=True),
                    "Cant_Diferencia": st.column_config.NumberColumn("Dif", step=1)
                },
                use_container_width=True,
                height=600,
                key=f"pc_{seleccion}"
            )
            if st.button("💾 Guardar PC", type="primary"):
                df_lote.update(editor)
                clean = df_master[df_master['Nombre_Lote'] != seleccion]
                final = pd.concat([clean, df_lote], ignore_index=True)
                conn.update(worksheet="Hoja1", data=final)
                st.success("Guardado.")

# ---------------------------------------------------------
# C. BORRAR
# ---------------------------------------------------------
elif accion == "Borrar Stock":
    lotes = [x for x in df_master['Nombre_Lote'].unique() if str(x) != 'nan']
    if lotes:
        b = st.selectbox("Borrar:", lotes)
        if st.button("🔥 Confirmar"):
            new = df_master[df_master['Nombre_Lote'] != b]
            conn.update(worksheet="Hoja1", data=new)
            st.success("Borrado."); time.sleep(1); st.rerun()

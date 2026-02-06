import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime
import time
import streamlit.components.v1 as components

# ==========================================
# 0. CONFIGURACIÓN Y WAKE LOCK (Pantalla Encendida)
# ==========================================
st.set_page_config(page_title="Stock Equus", layout="wide")

# Script para que no se apague la pantalla del celular
def keep_screen_awake():
    keep_awake_code = """
    <script>
    async function requestWakeLock() {
      try {
        const wakeLock = await navigator.wakeLock.request('screen');
        console.log('Pantalla mantenida encendida');
      } catch (err) {}
    }
    requestWakeLock();
    document.addEventListener('visibilitychange', async () => {
      if (document.visibilityState === 'visible') await requestWakeLock();
    });
    </script>
    """
    components.html(keep_awake_code, height=0)

keep_screen_awake()

# CSS para maximizar espacio en celular
st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .stButton > button {
        height: 3em; 
        font-weight: bold;
        font-size: 16px;
    }
    /* Ajuste para que los filtros no ocupen tanto lugar */
    .stExpander {
        border: 1px solid #ddd;
        border-radius: 8px;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 1. GESTIÓN DE SESIÓN
# ==========================================
if 'admin_logged_in' not in st.session_state: st.session_state.admin_logged_in = False

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
# 2. FUNCIONES DE LECTURA Y LIMPIEZA
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
                        if len(temp.columns)>1: df = temp; break
                    except: continue
                if df is not None: break
        
        if df is not None: return clean_columns(df)
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

def color_rows(val):
    if '✅ OK' in str(val): return 'background-color: #d4edda'
    if '🔴 Falta' in str(val): return 'background-color: #f8d7da'
    if '🟣 Sobra' in str(val): return 'background-color: #e2d9f3'
    return ''

# ==========================================
# 3. CONEXIÓN
# ==========================================
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data():
    try:
        df = conn.read(worksheet="Hoja1", ttl=0)
        req = ['Nombre_Lote', 'Estado_Stock', 'Código', 'Color', 'Stock 1ra un.', 'Descripción', 'Modelo_Ref', 'Cant_Diferencia']
        for c in req: 
            if c not in df.columns: df[c] = None
        return df
    except: return pd.DataFrame(columns=['Nombre_Lote'])

df_master = load_data()

# ==========================================
# 4. INTERFAZ SIMPLE
# ==========================================
st.sidebar.title("Menú")

if st.session_state.admin_logged_in:
    st.sidebar.success("Admin Activo")
    if st.sidebar.button("Salir"): logout(); st.rerun()
    accion = st.sidebar.radio("Acción:", ["Trabajar Stock", "Nuevo Stock", "Borrar Stock"])
else:
    with st.sidebar.expander("Admin Login"):
        st.text_input("Clave", type="password", key="password_input", on_change=check_admin_password)
    accion = st.sidebar.radio("Acción:", ["Trabajar Stock", "Nuevo Stock"])

st.title("👕 Control de Stock")

# ---------------------------------------------------------
# A. TRABAJAR STOCK (LO MÁS IMPORTANTE)
# ---------------------------------------------------------
if accion == "Trabajar Stock":
    lotes = [x for x in df_master['Nombre_Lote'].unique() if str(x) != 'nan']
    if not lotes:
        st.info("No hay stocks activos. Crea uno nuevo.")
    else:
        # Selección de Lote
        col_lote, col_btn = st.columns([3, 1])
        with col_lote:
            seleccion = st.selectbox("Selecciona Lote:", lotes)
        with col_btn:
            st.write("")
            if st.button("🔄", help="Recargar datos"): st.cache_data.clear(); st.rerun()

        # Cargar Datos
        df_lote = df_master[df_master['Nombre_Lote'] == seleccion].copy()
        
        # --- FILTROS POTENTES ---
        with st.expander("🔍 FILTROS (Toca para desplegar)", expanded=True):
            c1, c2, c3 = st.columns(3)
            
            # 1. Filtro Texto (Código)
            with c1:
                f_cod = st.text_input("Buscar Cód:", placeholder="...").strip()
            
            # 2. Filtro Descripción
            with c2:
                all_desc = sorted(df_lote['Descripción'].dropna().astype(str).unique())
                f_desc = st.selectbox("Descripción:", ["Todos"] + all_desc)
            
            # 3. Filtro Color (Dependiente)
            with c3:
                if f_desc != "Todos":
                    sub_df = df_lote[df_lote['Descripción'] == f_desc]
                    all_col = sorted(sub_df['Color'].dropna().astype(str).unique())
                else:
                    all_col = sorted(df_lote['Color'].dropna().astype(str).unique())
                f_col = st.selectbox("Color:", ["Todos"] + all_col)

        # APLICAR FILTROS
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

        st.caption(f"Mostrando **{len(df_view)}** items filtrados.")

        # --- TABLA DE EDICIÓN (ESTABLE) ---
        # Configuramos para que sea lo más amigable posible en móvil
        editor = st.data_editor(
            df_view[['Talle_Code', 'Color', 'Descripción', 'Stock 1ra un.', 'Estado_Stock', 'Cant_Diferencia']],
            column_config={
                "Talle_Code": st.column_config.TextColumn("Talle", disabled=True),
                "Color": st.column_config.TextColumn("Color", disabled=True),
                "Descripción": st.column_config.TextColumn("Desc.", disabled=True),
                "Stock 1ra un.": st.column_config.NumberColumn("Stk", disabled=True),
                "Estado_Stock": st.column_config.SelectboxColumn(
                    "Estado", 
                    options=["Pendiente", "✅ OK", "🔴 Falta", "🟣 Sobra"],
                    required=True,
                    width="small"
                ),
                "Cant_Diferencia": st.column_config.NumberColumn(
                    "Dif", 
                    min_value=-100, max_value=100, step=1,
                    width="small",
                    help="Usa botones +/- si aparecen"
                )
            },
            use_container_width=True,
            height=500, # Altura fija para evitar saltos
            hide_index=True,
            key=f"editor_final_{seleccion}"
        )
        
        # BOTÓN DE GUARDADO GIGANTE
        if st.button("💾 GUARDAR CAMBIOS EN LA NUBE", type="primary", use_container_width=True):
            # 1. Actualizar el lote local con lo editado
            df_lote.update(editor)
            
            # 2. Reconstruir el master (quitando lo viejo y poniendo lo nuevo)
            clean_master = df_master[df_master['Nombre_Lote'] != seleccion]
            final_upload = pd.concat([clean_master, df_lote], ignore_index=True)
            
            # 3. Subir
            conn.update(worksheet="Hoja1", data=final_upload)
            st.success("✅ Guardado exitoso!")
            time.sleep(1)
            st.rerun()
            
        st.divider()
        # Resumen rápido de diferencias
        dif = df_lote[df_lote['Estado_Stock'].isin(['🔴 Falta', '🟣 Sobra'])]
        if not dif.empty:
            st.warning(f"Resumen: Hay {len(dif)} items con diferencias.")
            st.dataframe(dif[['Descripción', 'Talle_Code', 'Estado_Stock', 'Cant_Diferencia']], use_container_width=True, hide_index=True)

# ---------------------------------------------------------
# B. NUEVO STOCK
# ---------------------------------------------------------
elif accion == "Nuevo Stock":
    st.header("Crear Nuevo Lote")
    lote = st.text_input("Nombre Sector")
    file = st.file_uploader("Archivo Excel/CSV")
    
    if st.button("🚀 Crear", type="primary", use_container_width=True) and file and lote:
        lote = lote.strip()
        if lote in [str(x) for x in df_master['Nombre_Lote'].unique()]:
            st.error("Nombre repetido.")
        else:
            df = load_universal_file(file)
            if df is not None and 'Código' in df.columns:
                df = sanitize_dataframe(df)
                parsed = df['Código'].apply(parse_product_code)
                df['Modelo_Ref'] = [x[0] for x in parsed]
                df['Color_Code'] = [x[1] for x in parsed]
                df['Talle_Code'] = [x[2] for x in parsed]
                df['Nombre_Lote'] = lote
                df['Estado_Stock'] = "Pendiente"
                df['Cant_Diferencia'] = 0
                for c in ['Color','Descripción']: 
                    if c not in df.columns: df[c] = "-"
                if 'Stock 1ra un.' not in df.columns: df['Stock 1ra un.'] = 0
                
                df = df.dropna(axis=1, how='all')
                master = pd.concat([df_master, df], ignore_index=True)
                conn.update(worksheet="Hoja1", data=master)
                st.success("Creado!"); time.sleep(1); st.rerun()
            else: st.error("Error en archivo.")

# ---------------------------------------------------------
# C. BORRAR (ADMIN)
# ---------------------------------------------------------
elif accion == "Borrar Stock":
    st.header("Zona Admin")
    lotes = [x for x in df_master['Nombre_Lote'].unique() if str(x) != 'nan']
    if lotes:
        b = st.selectbox("Borrar:", lotes)
        if st.button("🔥 Confirmar Borrado"):
            new = df_master[df_master['Nombre_Lote'] != b]
            conn.update(worksheet="Hoja1", data=new)
            st.success("Borrado."); time.sleep(1); st.rerun()

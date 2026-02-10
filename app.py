import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime
import time
import io

# Configuración básica
st.set_page_config(page_title="Stock Equus", layout="wide")

# CSS para ocultar elementos molestos y agrandar botones en el celu
st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    /* Botones más grandes para dedos */
    .stButton > button {
        height: 3em; 
        font-weight: bold;
    }
    /* Radio buttons más grandes */
    .stRadio label {
        font-size: 20px !important;
        padding: 10px;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 1. GESTIÓN DE SESIÓN
# ==========================================
if 'admin_logged_in' not in st.session_state:
    st.session_state.admin_logged_in = False

def check_admin_password():
    try:
        if st.session_state.password_input == st.secrets["admin"]["password"]:
            st.session_state.admin_logged_in = True
            st.session_state.password_input = "" 
        else:
            time.sleep(1)
            st.error("❌ Incorrecto")
    except:
        st.error("⚠️ Configura los secrets.")

def logout():
    st.session_state.admin_logged_in = False

# ==========================================
# 2. FUNCIONES
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
        # Intento genérico para Excel y CSV
        if filename.endswith(('.xls', '.xlsx')):
            try:
                df = pd.read_excel(uploaded_file)
            except:
                pass # Puede ser CSV disfrazado
        
        if df is None: # Probamos como CSV
            encodings = ['utf-8-sig', 'latin-1', 'cp1252']
            for enc in encodings:
                for sep in [',', ';', '\t']:
                    try:
                        uploaded_file.seek(0)
                        df_temp = pd.read_csv(uploaded_file, encoding=enc, sep=sep, engine='python')
                        if len(df_temp.columns) > 1:
                            df = df_temp
                            break
                    except: continue
                if df is not None: break
        
        if df is not None:
            # Buscar header
            df = clean_columns(df)
            if 'Código' not in df.columns:
                # Búsqueda manual de fila header
                for i in range(20):
                    # Lógica simplificada de header hunting
                    pass 
            return df
    except: return None
    return df

def parse_product_code(code_str):
    try:
        if pd.isna(code_str): return "S/D", "S/D", "S/D"
        parts = str(code_str).strip().split('.')
        if len(parts) < 2: return code_str, "N/A", "N/A"
        return parts[0], parts[1][:3] if len(parts[1])>=3 else parts[1], parts[1][-3:] if len(parts[1])>=6 else "N/A"
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
        expected = ['Nombre_Lote', 'Estado_Stock', 'Código', 'Color', 'Stock 1ra un.', 'Descripción', 'Modelo_Ref', 'Cant_Diferencia']
        for c in expected:
            if c not in df.columns: df[c] = None
        return df
    except:
        return pd.DataFrame(columns=['Nombre_Lote'])

df_master = load_data()

# ==========================================
# 4. INTERFAZ OPTIMIZADA
# ==========================================
st.sidebar.title("Menú")

# SELECTOR DE MODO (CRÍTICO PARA CELULAR)
modo_visual = st.sidebar.radio(
    "Dispositivo:", 
    ["📱 MODO CELULAR", "💻 MODO PC"],
    index=0,
    help="El modo celular usa botones grandes y bloquea el teclado."
)

st.sidebar.divider()

if st.session_state.admin_logged_in:
    st.sidebar.success("Admin Activo")
    if st.sidebar.button("Salir Admin"):
        logout()
        st.rerun()
    accion = st.sidebar.radio("Acción:", ["Trabajar Stock", "Nuevo Stock", "Borrar Stock"])
else:
    with st.sidebar.expander("Admin Login"):
        st.text_input("Clave", type="password", key="password_input", on_change=check_admin_password)
    accion = st.sidebar.radio("Acción:", ["Trabajar Stock", "Nuevo Stock"])

st.title("👕 Control de Stock")

# ---------------------------------------------------------
# A. NUEVO STOCK (Igual para ambos)
# ---------------------------------------------------------
if accion == "Nuevo Stock":
    st.header("Crear Nuevo Lote")
    lote = st.text_input("Nombre Sector")
    file = st.file_uploader("Archivo")
    if st.button("Procesar", type="primary", use_container_width=True) and file and lote:
        lote = lote.strip()
        if lote in [str(x) for x in df_master['Nombre_Lote'].unique()]:
            st.error("Ya existe.")
        else:
            df = load_universal_file(file)
            if df is not None:
                df = clean_columns(df)
                if 'Código' in df.columns:
                    df = sanitize_dataframe(df)
                    parsed = df['Código'].apply(parse_product_code)
                    df['Modelo_Ref'] = [x[0] for x in parsed]
                    df['Color_Code'] = [x[1] for x in parsed]
                    df['Talle_Code'] = [x[2] for x in parsed]
                    df['Nombre_Lote'] = lote
                    df['Estado_Stock'] = "Pendiente"
                    df['Cant_Diferencia'] = 0
                    
                    # Rellenar faltantes
                    for c in ['Color', 'Descripción']: 
                        if c not in df.columns: df[c] = "-"
                    if 'Stock 1ra un.' not in df.columns: df['Stock 1ra un.'] = 0
                    
                    df = df.dropna(axis=1, how='all')
                    master = pd.concat([df_master, df], ignore_index=True)
                    conn.update(worksheet="Hoja1", data=master)
                    st.success("Creado!")
                    time.sleep(1)
                    st.rerun()
            else:
                st.error("Error leyendo archivo.")

# ---------------------------------------------------------
# B. TRABAJAR STOCK
# ---------------------------------------------------------
elif accion == "Trabajar Stock":
    lotes = [x for x in df_master['Nombre_Lote'].unique() if str(x) != 'nan']
    if not lotes:
        st.info("No hay stocks.")
    else:
        seleccion = st.selectbox("Lote:", lotes)
        
        if st.button("🔄 Actualizar", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

        df_lote = df_master[df_master['Nombre_Lote'] == seleccion].copy()

        # --- FILTROS ---
        with st.expander("🔍 Filtros"):
            desc = st.multiselect("Descripción:", sorted(df_lote['Descripción'].astype(str).unique()))
            if desc:
                df_view = df_lote[df_lote['Descripción'].isin(desc)]
            else:
                df_view = df_lote
        
        if 'Modelo_Ref' in df_view.columns:
            df_view = df_view.sort_values(by=['Modelo_Ref', 'Color_Code', 'Talle_Code'])

        # =========================================================
        # MODO CELULAR (FICHA GIGANTE - SIN TECLADO)
        # =========================================================
        if modo_visual == "📱 MODO CELULAR":
            if 'idx' not in st.session_state: st.session_state.idx = 0
            if st.session_state.idx >= len(df_view): st.session_state.idx = 0
            
            if len(df_view) > 0:
                # Obtener Item Actual
                row = df_view.iloc[st.session_state.idx]
                real_idx = row.name # Índice real en el DF original
                
                # --- TARJETA VISUAL ---
                with st.container(border=True):
                    # Encabezado Grande
                    st.markdown(f"### {row['Descripción']}")
                    c1, c2 = st.columns(2)
                    c1.metric("Talle", str(row.get('Talle_Code','-')))
                    c2.metric("Color", str(row.get('Color','-')))
                    st.caption(f"Modelo: {row.get('Modelo_Ref','-')} | Código: {row.get('Código','-')}")
                    
                    st.divider()
                    
                    # 1. ESTADO (Radio Button Horizontal - NO ABRE TECLADO)
                    st.subheader("Estado:")
                    estados = ["Pendiente", "✅ OK", "🔴 Falta", "🟣 Sobra"]
                    curr_est = row['Estado_Stock'] if row['Estado_Stock'] in estados else "Pendiente"
                    
                    # Callback para guardar estado al toque
                    def update_estado():
                        # Esta función se ejecuta al cambiar el radio
                        pass 

                    # Usamos key dinámica para que se resetee al cambiar de producto
                    nuevo_estado = st.radio(
                        "Selecciona:", 
                        estados, 
                        index=estados.index(curr_est), 
                        horizontal=True,
                        key=f"rad_{st.session_state.idx}",
                        label_visibility="collapsed"
                    )

                    # 2. DIFERENCIA (Botones Grandes - NO ABRE TECLADO)
                    dif = float(row.get('Cant_Diferencia', 0))
                    
                    if nuevo_estado in ["🔴 Falta", "🟣 Sobra"]:
                        st.subheader("Diferencia:")
                        col_men, col_num, col_mas = st.columns([1, 1, 1])
                        
                        # Botones que actualizan session state temporal
                        if col_men.button("➖", key=f"d_m_{st.session_state.idx}", use_container_width=True):
                            dif -= 1
                        
                        col_num.markdown(f"<h1 style='text-align: center; margin: 0;'>{int(dif)}</h1>", unsafe_allow_html=True)
                        
                        if col_mas.button("➕", key=f"d_p_{st.session_state.idx}", use_container_width=True):
                            dif += 1
                    else:
                        dif = 0
                
                # --- NAVEGACIÓN Y GUARDADO ---
                c_ant, c_sig = st.columns(2)
                
                # Al navegar, guardamos en memoria (df_lote)
                # OJO: Los botones de dif ya actualizaron la variable 'dif' local, 
                # pero necesitamos persistirla antes de cambiar de índice.
                
                # Actualizamos el DF local con lo que hay en pantalla
                df_lote.at[real_idx, 'Estado_Stock'] = nuevo_estado
                df_lote.at[real_idx, 'Cant_Diferencia'] = dif
                
                if c_ant.button("⬅️ Anterior", use_container_width=True):
                    st.session_state.idx = max(0, st.session_state.idx - 1)
                    st.rerun()
                
                if c_sig.button("Siguiente ➡️", use_container_width=True):
                    st.session_state.idx = min(len(df_view) - 1, st.session_state.idx + 1)
                    st.rerun()

                st.write("")
                # GUARDAR EN LA NUBE (Botón Gigante)
                if st.button("💾 GUARDAR TODO EN LA NUBE", type="primary", use_container_width=True):
                    # Reconstruir master
                    clean = df_master[df_master['Nombre_Lote'] != seleccion]
                    final = pd.concat([clean, df_lote], ignore_index=True)
                    conn.update(worksheet="Hoja1", data=final)
                    st.success("Guardado!")
                
                st.caption(f"Producto {st.session_state.idx + 1} de {len(df_view)}")

            else:
                st.warning("No hay productos con ese filtro.")

        # =========================================================
        # MODO PC (TABLA DE SIEMPRE)
        # =========================================================
        else:
            editor = st.data_editor(
                df_view[['Talle_Code', 'Color', 'Descripción', 'Stock 1ra un.', 'Estado_Stock', 'Cant_Diferencia']],
                column_config={
                    "Estado_Stock": st.column_config.SelectboxColumn("Estado", options=["Pendiente", "✅ OK", "🔴 Falta", "🟣 Sobra"], required=True),
                    "Cant_Diferencia": st.column_config.NumberColumn("Dif", step=1)
                },
                use_container_width=True,
                height=600,
                key=f"edit_{seleccion}"
            )
            
            if st.button("💾 Guardar Cambios PC", type="primary"):
                df_lote.update(editor)
                clean = df_master[df_master['Nombre_Lote'] != seleccion]
                final = pd.concat([clean, df_lote], ignore_index=True)
                conn.update(worksheet="Hoja1", data=final)
                st.success("Guardado.")

# ---------------------------------------------------------
# C. BORRAR (ADMIN)
# ---------------------------------------------------------
elif accion == "Borrar Stock":
    st.header("Borrar Lote")
    lotes = [x for x in df_master['Nombre_Lote'].unique() if str(x) != 'nan']
    if lotes:
        b = st.selectbox("Elegir:", lotes)
        if st.button("🔥 BORRAR DEFINITIVAMENTE"):
            new = df_master[df_master['Nombre_Lote'] != b]
            conn.update(worksheet="Hoja1", data=new)
            st.success("Chau lote.")
            time.sleep(1)
            st.rerun()

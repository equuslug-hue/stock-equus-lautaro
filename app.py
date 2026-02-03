import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime
import time
import io

# Configuración "wide" ayuda, pero en celular se adapta solo
st.set_page_config(page_title="Stock Lautaro App", layout="wide")

# ==========================================
# 1. GESTIÓN DE SESIÓN Y CSS PARA MÓVIL
# ==========================================
if 'admin_logged_in' not in st.session_state:
    st.session_state.admin_logged_in = False

# Inyectamos CSS para ocultar elementos molestos en móvil y agrandar botones
st.markdown("""
<style>
    /* Ocultar el menú de hamburguesa de Streamlit y el footer para ganar espacio */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* Agrandar selectores para dedos */
    .stSelectbox div[data-baseweb="select"] > div {
        min-height: 50px;
    }
    
    /* Ajuste para tablas en móvil */
    .stDataFrame { font-size: 14px; }
</style>
""", unsafe_allow_html=True)

def check_admin_password():
    try:
        if st.session_state.password_input == st.secrets["admin"]["password"]:
            st.session_state.admin_logged_in = True
            st.session_state.password_input = "" 
        else:
            time.sleep(1) # Pequeño delay anti-fuerza bruta
            st.error("❌ Contraseña incorrecta")
    except Exception:
        st.error("⚠️ Error: Configura [admin] password en Secrets.")

def logout():
    st.session_state.admin_logged_in = False

# ==========================================
# 2. FUNCIONES DE LECTURA Y LIMPIEZA
# ==========================================
def clean_columns(df):
    """Normaliza nombres de columnas"""
    df.columns = [str(c).strip().replace('\ufeff', '') for c in df.columns]
    rename_map = {}
    for col in df.columns:
        c_lower = col.lower()
        if 'codigo' in c_lower or 'código' in c_lower: rename_map[col] = 'Código'
        elif 'descripcion' in c_lower or 'descripción' in c_lower: rename_map[col] = 'Descripción'
        elif 'stock' in c_lower and '1' in c_lower: rename_map[col] = 'Stock 1ra un.'
        elif c_lower == 'stock': rename_map[col] = 'Stock 1ra un.'
        elif 'precio' in c_lower: rename_map[col] = 'Precio'
            
    df.rename(columns=rename_map, inplace=True)
    return df

def try_read_csv_strategies(file_obj, strategies):
    for enc in strategies['encodings']:
        for sep in strategies['separators']:
            try:
                file_obj.seek(0)
                header_idx = 0
                found = False
                for i in range(30):
                    try:
                        line = file_obj.readline().decode(enc).lower()
                        if 'código' in line or 'codigo' in line:
                            header_idx = i
                            found = True
                            break
                    except: continue
                if not found: header_idx = 0 
                file_obj.seek(0)
                df = pd.read_csv(file_obj, header=header_idx, encoding=enc, sep=sep, engine='python')
                df = clean_columns(df)
                if 'Código' in df.columns: return df
            except: continue
    return None

def load_universal_file(uploaded_file):
    filename = uploaded_file.name.lower()
    df = None
    # HTML/XLS
    if filename.endswith('.xls'):
        try:
            uploaded_file.seek(0)
            dfs = pd.read_html(uploaded_file, header=0, encoding='latin-1')
            for d in dfs:
                d = clean_columns(d)
                if 'Código' in d.columns: return d
        except: pass
    # Excel Nativo
    if filename.endswith('.xlsx') or filename.endswith('.xls'):
        try:
            df_raw = pd.read_excel(uploaded_file, header=None)
            header_idx = -1
            for i, row in df_raw.head(20).iterrows():
                row_str = row.astype(str).str.lower().tolist()
                if any('código' in s or 'codigo' in s for s in row_str):
                    header_idx = i
                    break
            uploaded_file.seek(0)
            df = pd.read_excel(uploaded_file, header=header_idx if header_idx != -1 else 0)
            df = clean_columns(df)
            if 'Código' in df.columns: return df
        except: pass
    # CSV
    encodings = ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
    separators = [',', ';', '\t']
    return try_read_csv_strategies(uploaded_file, {'encodings': encodings, 'separators': separators})

def parse_product_code(code_str):
    try:
        if pd.isna(code_str): return "S/D", "S/D", "S/D"
        parts = str(code_str).strip().split('.')
        if len(parts) < 2: return code_str, "N/A", "N/A"
        modelo = parts[0]
        suffix = parts[1]
        color = suffix[:3] if len(suffix) >= 6 else suffix
        talle = suffix[-3:] if len(suffix) >= 6 else "N/A"
        return modelo, color, talle
    except: return str(code_str), "Err", "Err"

def sanitize_dataframe(df):
    """Sanitiza contra inyección de fórmulas"""
    def clean_cell(cell):
        if isinstance(cell, str) and len(cell) > 0:
            if cell.strip().startswith(('=', '+', '-', '@')): return f"'{cell}" 
        return cell
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].apply(clean_cell)
    return df

def color_rows(val):
    if '✅ OK' in str(val): return 'background-color: #d4edda; color: #155724'
    if '🔴 Falta' in str(val): return 'background-color: #f8d7da; color: #721c24'
    if '🟣 Sobra' in str(val): return 'background-color: #e2d9f3; color: #4a148c'
    return ''

# ==========================================
# 3. CONEXIÓN GOOGLE SHEETS
# ==========================================
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data():
    try:
        df = conn.read(worksheet="Hoja1", ttl=0)
        expected_cols = ['Nombre_Lote', 'Estado_Stock', 'Código', 'Color', 'Stock 1ra un.', 'Descripción', 'Modelo_Ref']
        for col in expected_cols:
            if col not in df.columns: df[col] = None 
        return df
    except:
        return pd.DataFrame(columns=['Nombre_Lote', 'Estado_Stock', 'Código'])

df_master = load_data()

# ==========================================
# 4. INTERFAZ
# ==========================================
st.sidebar.title("Menú")
opciones_menu = ["📂 Continuar Stock", "➕ Iniciar NUEVO"]
if st.session_state.admin_logged_in:
    opciones_menu.append("🛑 Zona Admin")
    st.sidebar.success("Modo Admin")
    if st.sidebar.button("Salir Admin", use_container_width=True):
        logout()
        st.rerun()
else:
    with st.sidebar.expander("🔐 Admin"):
        st.text_input("Clave:", type="password", key="password_input", on_change=check_admin_password)

modo = st.sidebar.radio("Ir a:", opciones_menu)
st.title("👕 Control de Stock")

# ---------------------------------------------------------
# ZONA A: INICIAR NUEVO
# ---------------------------------------------------------
if modo == "➕ Iniciar NUEVO":
    st.header("Nuevo Lote")
    nuevo_lote = st.text_input("Nombre del Sector", placeholder="Ej: Deposito_A")
    archivo = st.file_uploader("Subir Excel", type=['xlsx', 'xls', 'csv'])

    if st.button("🚀 Crear Lote", type="primary", use_container_width=True) and archivo and nuevo_lote:
        # Validación tamaño (5MB)
        if archivo.size > 5 * 1024 * 1024:
            st.error("Archivo muy pesado (>5MB).")
            st.stop()
            
        nombre_limpio = nuevo_lote.strip()
        if nombre_limpio in [str(x).strip() for x in df_master['Nombre_Lote'].unique()]:
            st.error(f"Ya existe '{nombre_limpio}'.")
        else:
            df_new = load_universal_file(archivo)
            if df_new is None or 'Código' not in df_new.columns:
                st.error("Error leyendo archivo. ¿Tiene columna 'Código'?")
            else:
                try:
                    df_new = sanitize_dataframe(df_new)
                    parsed = df_new['Código'].apply(parse_product_code)
                    df_new['Modelo_Ref'] = [x[0] for x in parsed]
                    df_new['Color_Code'] = [x[1] for x in parsed]
                    df_new['Talle_Code'] = [x[2] for x in parsed]
                    
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
                    st.success("Creado con éxito.")
                    time.sleep(1)
                    st.rerun()
                except Exception as e: st.error(f"Error: {e}")

# ---------------------------------------------------------
# ZONA B: CONTINUAR / EDITAR (OPTIMIZADO PARA CELULAR)
# ---------------------------------------------------------
elif modo == "📂 Continuar Stock":
    lotes = [x for x in df_master['Nombre_Lote'].unique() if str(x) != 'nan']
    
    if not lotes:
        st.info("No hay stocks.")
    else:
        seleccion = st.selectbox("Selecciona Lote:", lotes)
        if st.button("🔄 Recargar Datos", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
        
        df_lote = df_master[df_master['Nombre_Lote'] == seleccion].copy()
        
        # --- FILTROS ---
        with st.expander("🔍 Filtros (Click aquí)", expanded=False):
            desc_opts = sorted(df_lote['Descripción'].dropna().astype(str).unique())
            sel_desc = st.multiselect("Descripción:", desc_opts)
            
            if sel_desc:
                mod_opts = sorted(df_lote[df_lote['Descripción'].isin(sel_desc)]['Modelo_Ref'].astype(str).unique())
            else:
                mod_opts = sorted(df_lote['Modelo_Ref'].astype(str).unique()) if 'Modelo_Ref' in df_lote.columns else []
            sel_mod = st.multiselect("Modelo:", mod_opts)

        # Aplicar filtros
        df_view = df_lote.copy()
        if sel_desc: df_view = df_view[df_view['Descripción'].isin(sel_desc)]
        if sel_mod: df_view = df_view[df_view['Modelo_Ref'].isin(sel_mod)]
        if 'Modelo_Ref' in df_view.columns:
            df_view = df_view.sort_values(by=['Modelo_Ref', 'Color_Code', 'Talle_Code'])

        # --- MODO DE VISUALIZACIÓN ---
        vista_tarjeta = st.toggle("📱 Activar Modo Ficha (Ideal Celular)", value=False)

        # ---------------- MODO FICHA (CELULAR) ----------------
        if vista_tarjeta:
            st.info("Modo Ficha: Edita uno por uno sin teclado.")
            
            # Usamos session_state para navegar entre productos
            if 'idx_ficha' not in st.session_state: st.session_state.idx_ficha = 0
            
            # Aseguramos que el índice sea válido
            if st.session_state.idx_ficha >= len(df_view): st.session_state.idx_ficha = 0
            
            # Obtener ítem actual
            if len(df_view) > 0:
                item_actual = df_view.iloc[st.session_state.idx_ficha]
                idx_original = df_view.index[st.session_state.idx_ficha] # Índice real en el DF
                
                # Tarjeta visual
                with st.container(border=True):
                    st.caption(f"Producto {st.session_state.idx_ficha + 1} de {len(df_view)}")
                    st.write(f"**{item_actual['Descripción']}**")
                    st.write(f"Modelo: {item_actual.get('Modelo_Ref', 'N/A')} | Talle: **{item_actual.get('Talle_Code', 'N/A')}**")
                    st.write(f"Color: {item_actual.get('Color', 'N/A')} ({item_actual.get('Color_Code', 'N/A')})")
                    st.metric("Stock Sistema", int(item_actual.get('Stock 1ra un.', 0)))
                    
                    st.divider()
                    
                    # Controles de ESTADO
                    st.write("Estado Actual:")
                    estado_vals = ["Pendiente", "✅ OK", "🔴 Falta", "🟣 Sobra"]
                    # Buscamos índice actual
                    try: idx_estado = estado_vals.index(item_actual['Estado_Stock'])
                    except: idx_estado = 0
                    
                    nuevo_estado = st.radio("Validación:", estado_vals, index=idx_estado, horizontal=True, label_visibility="collapsed")
                    
                    # Controles de DIFERENCIA (Solo si no es OK ni Pendiente)
                    nueva_dif = float(item_actual.get('Cant_Diferencia', 0))
                    if nuevo_estado in ["🔴 Falta", "🟣 Sobra"]:
                        st.write("Cantidad Diferencia (+/-):")
                        c1, c2, c3 = st.columns([1, 2, 1])
                        with c1: 
                            if st.button("➖", use_container_width=True): nueva_dif -= 1
                        with c2: 
                            st.markdown(f"<h3 style='text-align: center;'>{int(nueva_dif)}</h3>", unsafe_allow_html=True)
                        with c3: 
                            if st.button("➕", use_container_width=True): nueva_dif += 1
                    else:
                        nueva_dif = 0 # Reset si es OK
                
                # Botones de Navegación y Guardado
                col_nav1, col_nav2 = st.columns(2)
                with col_nav1:
                    if st.button("⬅️ Anterior", use_container_width=True):
                        st.session_state.idx_ficha = max(0, st.session_state.idx_ficha - 1)
                        st.rerun()
                with col_nav2:
                    if st.button("Siguiente ➡️", use_container_width=True):
                        # GUARDADO AUTOMÁTICO AL PASAR
                        df_lote.at[idx_original, 'Estado_Stock'] = nuevo_estado
                        df_lote.at[idx_original, 'Cant_Diferencia'] = nueva_dif
                        
                        # Actualizar en la nube (opcional, puede ser lento cada vez)
                        # Por velocidad, mejor guardamos en memoria y el usuario da "Guardar Todo" al final
                        # Pero para seguridad, actualizamos la "vista" en memoria
                        
                        st.session_state.idx_ficha = min(len(df_view) - 1, st.session_state.idx_ficha + 1)
                        st.rerun()
                
                # Botón Guardar Cambios Real
                st.write("")
                if st.button("💾 CONFIRMAR CAMBIOS (Actual item)", type="primary", use_container_width=True):
                    df_lote.at[idx_original, 'Estado_Stock'] = nuevo_estado
                    df_lote.at[idx_original, 'Cant_Diferencia'] = nueva_dif
                    
                    df_clean = df_master[df_master['Nombre_Lote'] != seleccion]
                    df_final = pd.concat([df_clean, df_lote], ignore_index=True)
                    conn.update(worksheet="Hoja1", data=df_final)
                    st.toast("✅ Guardado!", icon="💾")

            else:
                st.warning("No hay productos con los filtros actuales.")

        # ---------------- MODO TABLA (CLÁSICO OPTIMIZADO) ----------------
        else:
            cols_cfg = {
                "Código": st.column_config.TextColumn("Cód.", disabled=True),
                "Modelo_Ref": st.column_config.TextColumn("Mod.", disabled=True),
                "Color_Code": st.column_config.TextColumn("Col.C", disabled=True),
                "Talle_Code": st.column_config.TextColumn("Talle", disabled=True),
                "Color": st.column_config.TextColumn("Color", disabled=True),
                "Stock 1ra un.": st.column_config.NumberColumn("Stk", disabled=True),
                "Descripción": st.column_config.TextColumn("Desc.", disabled=True),
                "Estado_Stock": st.column_config.SelectboxColumn(
                    "Estado", 
                    options=["Pendiente", "✅ OK", "🔴 Falta", "🟣 Sobra"],
                    required=True,
                    width="small" # Fuerza a que sea compacto
                ),
                "Cant_Diferencia": st.column_config.NumberColumn(
                    "Dif (+/-)", 
                    min_value=-100, # Permite negativos
                    max_value=100,
                    step=1, # <--- ESTO PONE LAS FLECHITAS
                    format="%d",
                    width="small"
                )
            }
            
            cols_show = ['Talle_Code', 'Color', 'Descripción', 'Stock 1ra un.', 'Estado_Stock', 'Cant_Diferencia']
            cols_final = [c for c in cols_show if c in df_view.columns]

            edited_view = st.data_editor(
                df_view[cols_final],
                column_config=cols_cfg,
                use_container_width=True,
                height=500,
                hide_index=True,
                key=f"editor_{seleccion}"
            )
            
            if st.button("💾 GUARDAR CAMBIOS", type="primary", use_container_width=True):
                df_lote.update(edited_view)
                df_clean = df_master[df_master['Nombre_Lote'] != seleccion]
                df_final = pd.concat([df_clean, df_lote], ignore_index=True)
                conn.update(worksheet="Hoja1", data=df_final)
                st.toast("Guardado exitoso!", icon="✅")
                time.sleep(1)
                st.rerun()

        # Resumen
        st.divider()
        dif = df_lote[df_lote['Estado_Stock'].isin(['🔴 Falta', '🟣 Sobra'])]
        if not dif.empty:
            st.warning(f"Hay {len(dif)} diferencias reportadas.")
            st.dataframe(dif[['Descripción', 'Talle_Code', 'Estado_Stock', 'Cant_Diferencia']], use_container_width=True, hide_index=True)

# ---------------------------------------------------------
# ZONA C: ADMIN
# ---------------------------------------------------------
elif modo == "🛑 Zona Admin":
    st.header("Zona Admin")
    lotes = [x for x in df_master['Nombre_Lote'].unique() if str(x) != 'nan']
    if lotes:
        borrar = st.selectbox("Borrar lote:", lotes)
        if st.button("🔥 ELIMINAR", type="primary", use_container_width=True):
            df_new = df_master[df_master['Nombre_Lote'] != borrar]
            conn.update(worksheet="Hoja1", data=df_new)
            st.success("Eliminado.")
            time.sleep(1)
            st.rerun()

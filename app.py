import streamlit as st
import streamlit.components.v1 as components
from streamlit_js_eval import streamlit_js_eval
import os
import pandas as pd
from datetime import datetime
import re, math, hashlib

st.set_page_config(
    page_title="VegSP 🌱",
    page_icon="🌱",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
    html, body, [class*="css"] { font-family: 'Segoe UI', sans-serif; }

    .vegsp-header {
        background: linear-gradient(135deg, #2E7D32 0%, #66BB6A 100%);
        padding: 1.5rem 1.5rem 1.2rem 1.5rem;
        border-radius: 16px;
        color: white;
        text-align: center;
        margin-bottom: 1.2rem;
    }
    .vegsp-header h1 { font-size: 2.4rem; margin: 0; letter-spacing: 2px; }
    .vegsp-header p  { font-size: 1rem; margin: 0.3rem 0 0 0; opacity: 0.9; }

    .card-wrap {
        border-radius: 14px;
        padding: 0.8rem 1.2rem 0.6rem 1.2rem;
        margin-bottom: 0.9rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        border-left: 6px solid;
    }
    .card-vegano      { background:#F1F8F2; border-color:#2E7D32; }
    .card-vegetariano { background:#FFFDE7; border-color:#F9A825; }
    .card-opcoes      { background:#E3F2FD; border-color:#1565C0; }

    .badge {
        display: inline-block;
        padding: 2px 10px;
        border-radius: 20px;
        font-size: 0.75rem;
        font-weight: 600;
        margin-right: 6px;
    }
    .badge-vegano      { background:#2E7D32; color:white; }
    .badge-vegetariano { background:#F9A825; color:white; }
    .badge-opcoes      { background:#1565C0; color:white; }

    .status-aberto  { color:#2E7D32; font-weight:700; font-size:0.82rem; }
    .status-fechado { color:#c62828; font-weight:700; font-size:0.82rem; }
    .distancia      { color:#888; font-size:0.8rem; font-weight:500; margin-left:6px; }

    .card-info { font-size:0.85rem; color:#555; margin:0.1rem 0; line-height:1.5; }
    .card-info a { color:#1565C0; }
    .nome-estab { font-size:1.1rem; font-weight:700; color:#1a1a1a; margin:0.2rem 0 0.4rem 0; }

    .loc-pill {
        display:inline-block;
        background:#e8f5e9;
        border:1px solid #c8e6c9;
        border-radius:20px;
        padding:4px 14px;
        font-size:0.82rem;
        color:#2E7D32;
        font-weight:600;
        margin-bottom:0.8rem;
    }

    .legenda { display:flex; gap:1rem; flex-wrap:wrap; margin-bottom:0.8rem; }
    .leg-item { display:flex; align-items:center; gap:6px; font-size:0.82rem; font-weight:600; }
    .dot { width:12px; height:12px; border-radius:50%; display:inline-block; }
    .dot-v  { background:#2E7D32; }
    .dot-vl { background:#F9A825; }
    .dot-o  { background:#1565C0; }

    .contador { font-size:0.88rem; color:#666; margin-bottom:0.8rem; }

    .footer {
        text-align:center; margin-top:3rem; padding:1rem;
        border-top:1px solid #ddd; color:#888; font-size:0.82rem;
    }
    .footer a { color:#2E7D32; text-decoration:none; font-weight:600; }

    @media (max-width: 768px) {
        .vegsp-header h1 { font-size: 1.9rem; }
        .vegsp-header p  { font-size: 0.85rem; }
        .card-wrap { padding: 0.7rem 0.9rem 0.5rem 0.9rem; }
    }
</style>
""", unsafe_allow_html=True)

# ── Dados ─────────────────────────────────────────────────────────────
ARQUIVO = "VegSP_lista.xlsx"
ABAS = {
    "Veganos":                  ("vegano",      "🟢 Vegano"),
    "Vegetarianos (Ovo Lacto)": ("vegetariano", "🟡 Vegetariano"),
    "Com Opções (Vegano+Veg)":  ("opcoes",      "🔵 Com Opções"),
}

def _file_hash(path):
    h = hashlib.md5()
    try:
        with open(path, "rb") as f: h.update(f.read())
    except: pass
    return h.hexdigest()

@st.cache_data
def carregar_dados(file_hash: str):
    frames = []
    xl = pd.ExcelFile(ARQUIVO)
    for aba, (tipo, rotulo) in ABAS.items():
        if aba in xl.sheet_names:
            df = pd.read_excel(xl, sheet_name=aba, header=2)
            cols_base = ["nome","tipo_estab","culinaria","bairro","endereco",
                         "hora_abre","hora_fecha","dias","contato","obs"]
            if len(df.columns) >= 12:
                df.columns = cols_base + ["lat","lng"] + list(df.columns[12:])
            else:
                df.columns = cols_base[:len(df.columns)]
                df["lat"] = None; df["lng"] = None
            df = df.dropna(subset=["nome"])
            df = df[df["nome"].astype(str).str.strip() != ""]
            df["tipo"] = tipo; df["rotulo"] = rotulo
            frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

df = carregar_dados(_file_hash(ARQUIVO))

# ── Haversine ─────────────────────────────────────────────────────────
def haversine(lat1, lon1, lat2, lon2):
    R = 6371000
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlam = math.radians(lon2 - lon1)
    a = math.sin(dphi/2)**2 + math.cos(phi1)*math.cos(phi2)*math.sin(dlam/2)**2
    return R * 2 * math.atan2(math.sqrt(a), math.sqrt(1-a))

def fmt_dist(m):
    return f"📍 {int(m)} m" if m < 1000 else f"📍 {m/1000:.1f} km"

# ── Aberto agora ──────────────────────────────────────────────────────
DIAS_MAP = {"seg":0,"ter":1,"qua":2,"qui":3,"sex":4,"sáb":5,"sab":5,"dom":6}

def dias_func(texto):
    if not isinstance(texto, str): return set(range(7))
    t = texto.lower(); ativos = set()
    for m in re.finditer(r'(\w+)\s+a\s+(\w+)', t):
        ini = DIAS_MAP.get(m.group(1)[:3]); fim = DIAS_MAP.get(m.group(2)[:3])
        if ini is not None and fim is not None:
            if fim >= ini: ativos.update(range(ini, fim+1))
            else: ativos.update(range(ini,7)); ativos.update(range(0,fim+1))
    for a, i in DIAS_MAP.items():
        if a in t: ativos.add(i)
    return ativos if ativos else set(range(7))

def esta_aberto(row):
    try:
        agora = datetime.now(); hm = agora.hour*60 + agora.minute
        if agora.weekday() not in dias_func(row["dias"]): return False
        def to_min(v):
            if pd.isna(v): return None
            s = str(v).strip()
            if ":" in s: h,m = s.split(":"); return int(h)*60+int(m)
            try: return int(float(s))*60
            except: return None
        ab = to_min(row["hora_abre"]); fe = to_min(row["hora_fecha"])
        if ab is None or fe is None: return None
        if fe < ab: return hm >= ab or hm <= fe
        return ab <= hm <= fe
    except: return None

# ── Header ────────────────────────────────────────────────────────────
st.markdown("""
<div class="vegsp-header">
  <h1>🌱 VegSP</h1>
  <p>Guia de estabelecimentos veganos e vegetarianos em São Paulo e região</p>
</div>
""", unsafe_allow_html=True)

# ── Geolocalização AUTOMÁTICA ─────────────────────────────────────────
# JS roda na carga da página, sem botão, sem filtro
# Salva lat/lng nos query params e recarrega se ainda não tiver
# Componente de geolocalização embutido — cacheado para não recriar a cada rerun
# Localização cacheada no session_state para sobreviver a reruns
if "user_lat" not in st.session_state:
    st.session_state.user_lat = 0.0
    st.session_state.user_lng = 0.0

# Só pede localização enquanto não tiver
if st.session_state.user_lat == 0:
    # JS embutido com botão visível:
    # - Desktop/Android: tenta automático silencioso
    # - iOS: mostra botão (iOS exige gesto do usuário no mesmo frame JS)
    _loc = streamlit_js_eval(
        js_expressions="""
        await new Promise(resolve => {
            // Estilos
            const s = document.createElement('style');
            s.textContent = `
                body { margin:0; padding:4px 0; font-family:'Segoe UI',sans-serif; background:transparent; }
                #gbtn { background:#2E7D32; color:#fff; border:none; border-radius:20px;
                        padding:6px 16px; font-size:13px; cursor:pointer; }
                #gbtn:disabled { background:#81C784; cursor:default; }
                #gmsg { font-size:12px; color:#666; margin-left:8px; }
            `;
            document.head.appendChild(s);

            // Botão (fallback para iOS)
            const btn = document.createElement('button');
            btn.id = 'gbtn';
            btn.textContent = '📍 Usar minha localização';
            document.body.appendChild(btn);
            const msg = document.createElement('span');
            msg.id = 'gmsg';
            document.body.appendChild(msg);

            function enviar(pos) {
                document.body.innerHTML =
                    '<span style="color:#2E7D32;font-size:13px">✅ Localização obtida</span>';
                resolve({ lat: pos.coords.latitude, lng: pos.coords.longitude });
            }

            btn.onclick = function() {
                btn.disabled = true; msg.textContent = '⏳ Buscando...';
                navigator.geolocation.getCurrentPosition(
                    enviar,
                    function() { btn.disabled=false; msg.textContent='⚠️ Tente novamente.'; },
                    { enableHighAccuracy: true, timeout: 12000 }
                );
            };

            // Tentativa automática silenciosa (funciona no desktop e Android)
            if (navigator.geolocation) {
                navigator.geolocation.getCurrentPosition(
                    function(pos) { enviar(pos); },
                    function() { /* silencioso — iOS vai usar o botão */ },
                    { enableHighAccuracy: false, timeout: 3000, maximumAge: 120000 }
                );
            }
        })
        """,
        want_output=True,
        key="geo_loc",
        height=38
    )
    if _loc and isinstance(_loc, dict) and _loc.get("lat"):
        try:
            st.session_state.user_lat = float(_loc["lat"])
            st.session_state.user_lng = float(_loc["lng"])
        except:
            pass

user_lat = st.session_state.user_lat
user_lng = st.session_state.user_lng
tem_loc  = user_lat != 0

if tem_loc:
    st.markdown(
        '<span class="loc-pill">📍 Ordenando pelo mais próximo de você</span>',
        unsafe_allow_html=True
    )

# ── Filtros colapsáveis ───────────────────────────────────────────────
with st.expander("🔍 Filtros", expanded=False):
    col1, col2 = st.columns(2)
    with col1:
        aberto_agora = st.toggle("🕐 Aberto agora", value=False)
    with col2:
        busca_nome = st.text_input("🔎 Nome", placeholder="Ex: Salad Days")

    col3, col4 = st.columns(2)
    with col3:
        # Categoria como dropdown (multiselect igual aos outros)
        cat_opcoes = ["Todos", "🟢 Vegano", "🟡 Vegetariano", "🔵 Com Opções"]
        cat_sel = st.multiselect("Categoria", cat_opcoes[1:], default=cat_opcoes[1:])
    with col4:
        tipos_disp = sorted(df["tipo_estab"].dropna().unique().tolist()) if not df.empty else []
        tipo_sel = st.multiselect("Tipo de estabelecimento", tipos_disp)

    col5, col6 = st.columns(2)
    with col5:
        culin_disp = sorted(df["culinaria"].dropna().unique().tolist()) if not df.empty else []
        culin_sel = st.multiselect("Culinária", culin_disp)
    with col6:
        # Bairro com autocomplete via selectbox
        bairros_disp = ["(Todos os bairros)"] + sorted(
            df["bairro"].dropna().unique().tolist()) if not df.empty else ["(Todos os bairros)"]
        bairro_sel = st.selectbox("📍 Bairro ou cidade", bairros_disp,
                                  help="Digite para filtrar")

    aplicar = st.button("✅ Aplicar filtros", use_container_width=True)

# Session state para persistir filtros
if "filtros" not in st.session_state:
    st.session_state.filtros = {
        "aberto_agora": False,
        "cat_sel": ["🟢 Vegano", "🟡 Vegetariano", "🔵 Com Opções"],
        "tipo_sel": [], "culin_sel": [],
        "bairro_sel": "(Todos os bairros)", "busca_nome": "",
    }

if aplicar:
    st.session_state.filtros = {
        "aberto_agora": aberto_agora,
        "cat_sel":      cat_sel if cat_sel else ["🟢 Vegano", "🟡 Vegetariano", "🔵 Com Opções"],
        "tipo_sel":     tipo_sel,
        "culin_sel":    culin_sel,
        "bairro_sel":   bairro_sel,
        "busca_nome":   busca_nome,
    }

f = st.session_state.filtros

# ── Filtragem ─────────────────────────────────────────────────────────
resultado = df.copy() if not df.empty else pd.DataFrame()

if not resultado.empty:
    map_cat = {
        "🟢 Vegano":      "vegano",
        "🟡 Vegetariano": "vegetariano",
        "🔵 Com Opções":  "opcoes",
    }
    tipos_filtro = [map_cat[c] for c in f["cat_sel"] if c in map_cat]
    resultado = resultado[resultado["tipo"].isin(tipos_filtro)] if tipos_filtro else resultado.iloc[0:0]

    if f["tipo_sel"]:
        resultado = resultado[resultado["tipo_estab"].isin(f["tipo_sel"])]
    if f["culin_sel"]:
        resultado = resultado[resultado["culinaria"].isin(f["culin_sel"])]
    if f["bairro_sel"] != "(Todos os bairros)":
        resultado = resultado[resultado["bairro"] == f["bairro_sel"]]
    if f["busca_nome"].strip():
        resultado = resultado[resultado["nome"].str.contains(f["busca_nome"], case=False, na=False)]
    if f["aberto_agora"]:
        resultado["_ab"] = resultado.apply(esta_aberto, axis=1)
        resultado = resultado[resultado["_ab"] == True]

    # Calcular distância e ordenar automaticamente se tiver localização
    if tem_loc:
        def calc_dist(row):
            try:
                if pd.notna(row.get("lat")) and pd.notna(row.get("lng")):
                    return haversine(user_lat, user_lng, float(row["lat"]), float(row["lng"]))
            except: pass
            return None
        resultado["_dist"] = resultado.apply(calc_dist, axis=1)
        resultado = resultado.sort_values("_dist", na_position="last")
    else:
        resultado["_dist"] = None

# ── Legenda ───────────────────────────────────────────────────────────
st.markdown("""
<div class="legenda">
  <div class="leg-item"><span class="dot dot-v"></span> Vegano</div>
  <div class="leg-item"><span class="dot dot-vl"></span> Vegetariano (Ovo Lacto)</div>
  <div class="leg-item"><span class="dot dot-o"></span> Com Opções Veganas/Veg</div>
</div>
""", unsafe_allow_html=True)

# ── Resultados ────────────────────────────────────────────────────────
if resultado.empty:
    st.info("Nenhum estabelecimento encontrado. Ajuste os filtros! 🌱")
else:
    total = len(resultado)
    st.markdown(f'<p class="contador">Exibindo <strong>{total}</strong> estabelecimento{"s" if total!=1 else ""}</p>',
                unsafe_allow_html=True)

    for _, row in resultado.iterrows():
        tipo   = row["tipo"]
        rotulo = row["rotulo"]

        status = esta_aberto(row)
        if status is True:    sh = '<span class="status-aberto">🟢 Aberto agora</span>'
        elif status is False: sh = '<span class="status-fechado">🔴 Fechado agora</span>'
        else:                 sh = ""

        dv = row.get("_dist")
        dh = f'<span class="distancia">· {fmt_dist(dv)}</span>' if (dv is not None and not pd.isna(dv)) else ""

        hora_txt = "—"
        if pd.notna(row.get("hora_abre")) and pd.notna(row.get("hora_fecha")):
            hora_txt = f'{str(row["hora_abre"]).strip()} – {str(row["hora_fecha"]).strip()}'
        dias_txt = str(row["dias"]).strip() if pd.notna(row.get("dias")) else "—"

        contato = row.get("contato","")
        ch = ""
        if pd.notna(contato) and str(contato).strip():
            link = str(contato).strip()
            if not link.startswith("http"): link = "https://" + link
            ch = f'<p class="card-info">🔗 <a href="{link}" target="_blank">{str(contato).strip()}</a></p>'

        obs = row.get("obs","")
        oh = f'<p class="card-info">💬 {obs}</p>' if pd.notna(obs) and str(obs).strip() else ""

        st.markdown(f"""
        <div class="card-wrap card-{tipo}">
          <span class="badge badge-{tipo}">{rotulo}</span>{sh}
          <p class="nome-estab">{row['nome']}{dh}</p>
          <p class="card-info">🍽️ {row.get('tipo_estab','—')} &nbsp;|&nbsp; 🥘 {row.get('culinaria','—')}</p>
          <p class="card-info">📍 {row.get('bairro','—')}</p>
          <p class="card-info">🏠 {row.get('endereco','—')}</p>
          <p class="card-info">⏰ {hora_txt} &nbsp;|&nbsp; 📅 {dias_txt}</p>
          {ch}{oh}
        </div>
        """, unsafe_allow_html=True)

# ── Rodapé ────────────────────────────────────────────────────────────
st.markdown("""
<div class="footer">
  Feito com 🌱 por <a href="https://instagram.com/chrisporto80" target="_blank">@chrisporto80</a>
  &nbsp;|&nbsp; VegSP — Guia vegano e vegetariano de São Paulo
</div>
""", unsafe_allow_html=True)

# ── CÓDIGO DO GOOGLE ANALYTICS ──
components.html("""
<script async src="https://www.googletagmanager.com/gtag/js?id=G-KS41R8XPW2"></script>
<script>
  window.dataLayer = window.dataLayer || [];
  function gtag(){dataLayer.push(arguments);}
  gtag('js', new Date());

  gtag('config', 'G-KS41R8XPW2');
</script>
""", height=0, width=0)

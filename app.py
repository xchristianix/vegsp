import streamlit as st
import pandas as pd
from datetime import datetime
import re

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

    /* Card wrapper — só cor de fundo e borda, sem texto HTML dentro */
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

    .card-info { font-size:0.85rem; color:#555; margin: 0.1rem 0; line-height:1.5; }
    .card-info a { color:#1565C0; }

    /* Nome do estabelecimento — via st.markdown nativo, não HTML */
    .nome-estab {
        font-size: 1.1rem;
        font-weight: 700;
        color: #1a1a1a;
        margin: 0.2rem 0 0.5rem 0;
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

    /* Botão aplicar */
    div[data-testid="stButton"] > button {
        background-color: #2E7D32 !important;
        color: white !important;
        border: none !important;
        width: 100%;
        font-size: 1rem;
        padding: 0.55rem;
        border-radius: 10px;
    }

    /* Elimina padding extra no mobile */
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

import os, hashlib

def _file_hash(path):
    h = hashlib.md5()
    try:
        with open(path, "rb") as f:
            h.update(f.read())
    except:
        pass
    return h.hexdigest()

@st.cache_data
def carregar_dados(file_hash: str):
    # file_hash força o Streamlit a reler quando a planilha mudar
    frames = []
    xl = pd.ExcelFile(ARQUIVO)
    for aba, (tipo, rotulo) in ABAS.items():
        if aba in xl.sheet_names:
            df = pd.read_excel(xl, sheet_name=aba, header=2)
            df.columns = [
                "nome", "tipo_estab", "culinaria",
                "bairro", "endereco", "hora_abre", "hora_fecha",
                "dias", "contato", "obs"
            ]
            df = df.dropna(subset=["nome"])
            df = df[df["nome"].astype(str).str.strip() != ""]
            df["tipo"]   = tipo
            df["rotulo"] = rotulo
            frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

df = carregar_dados(_file_hash(ARQUIVO))

# ── Aberto agora ──────────────────────────────────────────────────────
DIAS_MAP = {"seg":0,"ter":1,"qua":2,"qui":3,"sex":4,"sáb":5,"sab":5,"dom":6}

def dias_funcionando(texto):
    if not isinstance(texto, str):
        return set(range(7))
    t = texto.lower()
    ativos = set()
    for m in re.finditer(r'(\w+)\s+a\s+(\w+)', t):
        ini = DIAS_MAP.get(m.group(1)[:3])
        fim = DIAS_MAP.get(m.group(2)[:3])
        if ini is not None and fim is not None:
            if fim >= ini:
                ativos.update(range(ini, fim+1))
            else:
                ativos.update(range(ini, 7))
                ativos.update(range(0, fim+1))
    for abrev, idx in DIAS_MAP.items():
        if abrev in t:
            ativos.add(idx)
    return ativos if ativos else set(range(7))

def esta_aberto(row):
    try:
        agora    = datetime.now()
        hora_min = agora.hour * 60 + agora.minute
        if agora.weekday() not in dias_funcionando(row["dias"]):
            return False
        def to_min(v):
            if pd.isna(v): return None
            s = str(v).strip()
            if ":" in s:
                h, m = s.split(":")
                return int(h)*60 + int(m)
            try: return int(float(s))*60
            except: return None
        abre  = to_min(row["hora_abre"])
        fecha = to_min(row["hora_fecha"])
        if abre is None or fecha is None: return None
        if fecha < abre:
            return hora_min >= abre or hora_min <= fecha
        return abre <= hora_min <= fecha
    except:
        return None

# ── Header ────────────────────────────────────────────────────────────
st.markdown("""
<div class="vegsp-header">
  <h1>🌱 VegSP</h1>
  <p>Guia de estabelecimentos veganos e vegetarianos em São Paulo e região</p>
</div>
""", unsafe_allow_html=True)

# ── Filtros colapsáveis ───────────────────────────────────────────────
with st.expander("🔍 Filtros", expanded=False):
    col1, col2 = st.columns(2)
    with col1:
        aberto_agora = st.toggle("🕐 Aberto agora", value=False)
    with col2:
        busca_nome = st.text_input("🔎 Nome", placeholder="Ex: Quintal Vegano")

    categorias = st.multiselect(
        "Categoria",
        ["🟢 Vegano", "🟡 Vegetariano", "🔵 Com Opções"],
        default=["🟢 Vegano", "🟡 Vegetariano", "🔵 Com Opções"],
    )

    col3, col4 = st.columns(2)
    with col3:
        tipos_disp = sorted(df["tipo_estab"].dropna().unique().tolist()) if not df.empty else []
        tipo_sel = st.multiselect("Tipo", tipos_disp)
    with col4:
        culin_disp = sorted(df["culinaria"].dropna().unique().tolist()) if not df.empty else []
        culin_sel = st.multiselect("Culinária", culin_disp)

    busca_bairro = st.text_input("📍 Bairro ou cidade", placeholder="Ex: Pinheiros, Campinas")

    aplicar = st.button("✅ Aplicar filtros")

# session_state para persistir filtros após fechar o expander
if "filtros" not in st.session_state:
    st.session_state.filtros = {
        "aberto_agora": False,
        "categorias": ["🟢 Vegano", "🟡 Vegetariano", "🔵 Com Opções"],
        "tipo_sel": [], "culin_sel": [], "busca_bairro": "", "busca_nome": "",
    }

if aplicar:
    st.session_state.filtros = {
        "aberto_agora": aberto_agora,
        "categorias":   categorias,
        "tipo_sel":     tipo_sel,
        "culin_sel":    culin_sel,
        "busca_bairro": busca_bairro,
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
    tipos_filtro = [map_cat[c] for c in f["categorias"] if c in map_cat]
    resultado = resultado[resultado["tipo"].isin(tipos_filtro)] if tipos_filtro else resultado.iloc[0:0]

    if f["tipo_sel"]:
        resultado = resultado[resultado["tipo_estab"].isin(f["tipo_sel"])]
    if f["culin_sel"]:
        resultado = resultado[resultado["culinaria"].isin(f["culin_sel"])]
    if f["busca_bairro"].strip():
        resultado = resultado[resultado["bairro"].str.contains(f["busca_bairro"], case=False, na=False)]
    if f["busca_nome"].strip():
        resultado = resultado[resultado["nome"].str.contains(f["busca_nome"], case=False, na=False)]
    if f["aberto_agora"]:
        resultado["_aberto"] = resultado.apply(esta_aberto, axis=1)
        resultado = resultado[resultado["_aberto"] == True]

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
    st.info("Nenhum estabelecimento encontrado. Ajuste os filtros ou adicione mais locais na planilha! 🌱")
else:
    total = len(resultado)
    s = "s" if total != 1 else ""
    st.markdown(f'<p class="contador">Exibindo <strong>{total}</strong> estabelecimento{s}</p>', unsafe_allow_html=True)

    for _, row in resultado.iterrows():
        tipo   = row["tipo"]
        rotulo = row["rotulo"]

        status = esta_aberto(row)
        status_html = ""
        if status is True:
            status_html = '<span class="status-aberto">🟢 Aberto agora</span>'
        elif status is False:
            status_html = '<span class="status-fechado">🔴 Fechado agora</span>'

        hora_txt = "—"
        if pd.notna(row.get("hora_abre")) and pd.notna(row.get("hora_fecha")):
            hora_txt = f'{str(row["hora_abre"]).strip()} – {str(row["hora_fecha"]).strip()}'
        dias_txt = str(row["dias"]).strip() if pd.notna(row.get("dias")) else "—"

        contato = row.get("contato", "")
        contato_html = ""
        if pd.notna(contato) and str(contato).strip():
            link = str(contato).strip()
            if not link.startswith("http"):
                link = "https://" + link
            contato_html = f'<p class="card-info">🔗 <a href="{link}" target="_blank">{str(contato).strip()}</a></p>'

        obs = row.get("obs", "")
        obs_html = f'<p class="card-info">💬 {obs}</p>' if pd.notna(obs) and str(obs).strip() else ""

        nome = str(row["nome"])
        tipo_estab = str(row.get("tipo_estab", "—"))
        culinaria  = str(row.get("culinaria",  "—"))
        bairro     = str(row.get("bairro",     "—"))
        endereco   = str(row.get("endereco",   "—"))

        # Card: div de fundo + conteúdo misto (HTML p/ badges, st.markdown p/ nome)
        st.markdown(f"""
        <div class="card-wrap card-{tipo}">
          <span class="badge badge-{tipo}">{rotulo}</span>{status_html}
          <p class="nome-estab">{nome}</p>
          <p class="card-info">🍽️ {tipo_estab} &nbsp;|&nbsp; 🥘 {culinaria}</p>
          <p class="card-info">📍 {bairro}</p>
          <p class="card-info">🏠 {endereco}</p>
          <p class="card-info">⏰ {hora_txt} &nbsp;|&nbsp; 📅 {dias_txt}</p>
          {contato_html}{obs_html}
        </div>
        """, unsafe_allow_html=True)

# ── Rodapé ────────────────────────────────────────────────────────────
st.markdown("""
<div class="footer">
  Feito com 🌱 por <a href="https://instagram.com/chrisporto80" target="_blank">@chrisporto80</a>
  &nbsp;|&nbsp; VegSP — Guia vegano e vegetariano de São Paulo
</div>
""", unsafe_allow_html=True)

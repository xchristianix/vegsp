import streamlit as st
import pandas as pd
from datetime import datetime
import re

# ── Configuração da página ────────────────────────────────────────────
st.set_page_config(
    page_title="VegSP 🌱",
    page_icon="🌱",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Estilos ───────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Fundo e fonte geral */
    html, body, [class*="css"] { font-family: 'Segoe UI', sans-serif; }

    /* Header principal */
    .vegsp-header {
        background: linear-gradient(135deg, #2E7D32 0%, #66BB6A 100%);
        padding: 2rem 2rem 1.5rem 2rem;
        border-radius: 16px;
        color: white;
        text-align: center;
        margin-bottom: 1.5rem;
    }
    .vegsp-header h1 { font-size: 2.8rem; margin: 0; letter-spacing: 2px; }
    .vegsp-header p  { font-size: 1.1rem; margin: 0.4rem 0 0 0; opacity: 0.9; }

    /* Cards de estabelecimento */
    .card {
        border-radius: 14px;
        padding: 1.2rem 1.4rem;
        margin-bottom: 1rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        border-left: 6px solid;
    }
    .card-vegano      { background:#F1F8F2; border-color:#2E7D32; }
    .card-vegetariano { background:#FFFDE7; border-color:#F9A825; }
    .card-opcoes      { background:#E3F2FD; border-color:#1565C0; }

    .card h3 { margin: 0 0 0.3rem 0; font-size: 1.15rem; }
    .card .badge {
        display: inline-block;
        padding: 2px 10px;
        border-radius: 20px;
        font-size: 0.75rem;
        font-weight: 600;
        margin-bottom: 0.5rem;
    }
    .badge-vegano      { background:#2E7D32; color:white; }
    .badge-vegetariano { background:#F9A825; color:white; }
    .badge-opcoes      { background:#1565C0; color:white; }

    .card .info { font-size: 0.88rem; color: #555; margin: 0.15rem 0; }
    .card .aberto { color: #2E7D32; font-weight: 700; }
    .card .fechado { color: #c62828; font-weight: 700; }

    /* Legenda */
    .legenda {
        display: flex; gap: 1rem; flex-wrap: wrap;
        margin-bottom: 1rem;
    }
    .leg-item {
        display: flex; align-items: center; gap: 6px;
        font-size: 0.85rem; font-weight: 600;
    }
    .dot { width:14px; height:14px; border-radius:50%; display:inline-block; }
    .dot-v  { background:#2E7D32; }
    .dot-vl { background:#F9A825; }
    .dot-o  { background:#1565C0; }

    /* Rodapé */
    .footer {
        text-align: center;
        margin-top: 3rem;
        padding: 1.2rem;
        border-top: 1px solid #ddd;
        color: #888;
        font-size: 0.85rem;
    }
    .footer a { color: #2E7D32; text-decoration: none; font-weight: 600; }

    /* Contador */
    .contador { font-size: 0.9rem; color:#666; margin-bottom: 1rem; }
</style>
""", unsafe_allow_html=True)

# ── Carrega dados ─────────────────────────────────────────────────────
ARQUIVO = "VegSP_lista.xlsx"
ABAS = {
    "Veganos":                        ("vegano",      "🟢 Vegano"),
    "Vegetarianos (Ovo Lacto)":       ("vegetariano", "🟡 Vegetariano"),
    "Com Opções (Vegano+Veg)":        ("opcoes",      "🔵 Com Opções"),
}

@st.cache_data
def carregar_dados():
    frames = []
    xl = pd.ExcelFile(ARQUIVO)
    for aba, (tipo, rotulo) in ABAS.items():
        if aba in xl.sheet_names:
            df = pd.read_excel(xl, sheet_name=aba, header=2)   # linha 3 = cabeçalho
            df.columns = [
                "nome", "tipo_estab", "culinaria",
                "bairro", "endereco", "hora_abre", "hora_fecha",
                "dias", "contato", "obs"
            ]
            df = df.dropna(subset=["nome"])
            df["tipo"] = tipo
            df["rotulo"] = rotulo
            frames.append(df)
    if frames:
        return pd.concat(frames, ignore_index=True)
    return pd.DataFrame()

df = carregar_dados()

# ── Função: está aberto agora? ────────────────────────────────────────
DIAS_MAP = {
    "seg": 0, "ter": 1, "qua": 2, "qui": 3,
    "sex": 4, "sáb": 5, "sab": 5, "dom": 6,
}

def dias_funcionando(texto_dias):
    """Retorna set de índices weekday (0=seg … 6=dom)."""
    if not isinstance(texto_dias, str):
        return set(range(7))
    texto = texto_dias.lower()
    ativos = set()
    # Intervalos como "seg a sex"
    for m in re.finditer(r'(\w+)\s+a\s+(\w+)', texto):
        ini = DIAS_MAP.get(m.group(1)[:3])
        fim = DIAS_MAP.get(m.group(2)[:3])
        if ini is not None and fim is not None:
            if fim >= ini:
                ativos.update(range(ini, fim + 1))
            else:
                ativos.update(range(ini, 7))
                ativos.update(range(0, fim + 1))
    # Dias avulsos
    for abrev, idx in DIAS_MAP.items():
        if abrev in texto:
            ativos.add(idx)
    return ativos if ativos else set(range(7))

def esta_aberto(row):
    try:
        agora = datetime.now()
        weekday = agora.weekday()
        hora_atual = agora.hour * 60 + agora.minute

        dias = dias_funcionando(row["dias"])
        if weekday not in dias:
            return False

        def to_min(val):
            if pd.isna(val):
                return None
            s = str(val).strip()
            if ":" in s:
                h, m = s.split(":")
                return int(h) * 60 + int(m)
            try:
                h = int(float(s))
                return h * 60
            except:
                return None

        abre  = to_min(row["hora_abre"])
        fecha = to_min(row["hora_fecha"])
        if abre is None or fecha is None:
            return None   # Sem info

        if fecha < abre:   # passa da meia-noite
            return hora_atual >= abre or hora_atual <= fecha
        return abre <= hora_atual <= fecha
    except:
        return None

# ── Header ────────────────────────────────────────────────────────────
st.markdown("""
<div class="vegsp-header">
  <h1>🌱 VegSP</h1>
  <p>Guia de estabelecimentos veganos e vegetarianos em São Paulo e região metropolitana</p>
</div>
""", unsafe_allow_html=True)

# ── Sidebar – Filtros ─────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🔍 Filtros")

    aberto_agora = st.toggle("🕐 Aberto agora", value=False)

    categorias = st.multiselect(
        "Categoria",
        ["🟢 Vegano", "🟡 Vegetariano", "🔵 Com Opções"],
        default=["🟢 Vegano", "🟡 Vegetariano", "🔵 Com Opções"],
    )

    tipos_disponiveis = sorted(df["tipo_estab"].dropna().unique()) if not df.empty else []
    tipo_sel = st.multiselect("Tipo de estabelecimento", tipos_disponiveis)

    culinarias_disp = sorted(df["culinaria"].dropna().unique()) if not df.empty else []
    culinaria_sel = st.multiselect("Tipo de culinária", culinarias_disp)

    busca_bairro = st.text_input("📍 Buscar por bairro ou cidade")

    busca_nome = st.text_input("🔎 Buscar por nome")

# ── Filtragem ─────────────────────────────────────────────────────────
resultado = df.copy() if not df.empty else pd.DataFrame()

if not resultado.empty:
    # Categoria
    map_cat = {
        "🟢 Vegano":      "vegano",
        "🟡 Vegetariano": "vegetariano",
        "🔵 Com Opções":  "opcoes",
    }
    tipos_filtro = [map_cat[c] for c in categorias if c in map_cat]
    resultado = resultado[resultado["tipo"].isin(tipos_filtro)]

    if tipo_sel:
        resultado = resultado[resultado["tipo_estab"].isin(tipo_sel)]
    if culinaria_sel:
        resultado = resultado[resultado["culinaria"].isin(culinaria_sel)]
    if busca_bairro:
        resultado = resultado[resultado["bairro"].str.contains(busca_bairro, case=False, na=False)]
    if busca_nome:
        resultado = resultado[resultado["nome"].str.contains(busca_nome, case=False, na=False)]
    if aberto_agora:
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
    st.info("Nenhum estabelecimento encontrado com os filtros selecionados.")
else:
    total = len(resultado)
    st.markdown(f'<p class="contador">Exibindo <strong>{total}</strong> estabelecimento{"s" if total != 1 else ""}</p>', unsafe_allow_html=True)

    for _, row in resultado.iterrows():
        tipo   = row["tipo"]
        rotulo = row["rotulo"]

        css_card  = f"card-{tipo}"
        css_badge = f"badge-{tipo}"

        # Status aberto/fechado
        status_aberto = esta_aberto(row)
        if aberto_agora:
            status_html = '<span class="aberto">🟢 Aberto agora</span>'
        elif status_aberto is True:
            status_html = '<span class="aberto">🟢 Aberto agora</span>'
        elif status_aberto is False:
            status_html = '<span class="fechado">🔴 Fechado agora</span>'
        else:
            status_html = ""

        dias_txt   = row["dias"] if pd.notna(row.get("dias")) else "—"
        hora_txt   = ""
        if pd.notna(row.get("hora_abre")) and pd.notna(row.get("hora_fecha")):
            hora_txt = f'{str(row["hora_abre"]).strip()} – {str(row["hora_fecha"]).strip()}'

        contato = row.get("contato", "")
        contato_html = ""
        if pd.notna(contato) and str(contato).strip():
            link = str(contato).strip()
            if not link.startswith("http"):
                link = "https://" + link
            contato_html = f'<p class="info">🔗 <a href="{link}" target="_blank">{str(contato).strip()}</a></p>'

        obs = row.get("obs", "")
        obs_html = f'<p class="info">💬 {obs}</p>' if pd.notna(obs) and str(obs).strip() else ""

        st.markdown(f"""
        <div class="card {css_card}">
          <span class="badge {css_badge}">{rotulo}</span>
          {status_html}
          <h3>{row['nome']}</h3>
          <p class="info">🍽️ {row.get('tipo_estab','—')} &nbsp;|&nbsp; 🥘 {row.get('culinaria','—')}</p>
          <p class="info">📍 {row.get('bairro','—')}</p>
          <p class="info">🏠 {row.get('endereco','—')}</p>
          <p class="info">⏰ {hora_txt} &nbsp;|&nbsp; 📅 {dias_txt}</p>
          {contato_html}
          {obs_html}
        </div>
        """, unsafe_allow_html=True)

# ── Rodapé ────────────────────────────────────────────────────────────
st.markdown("""
<div class="footer">
  Feito com 🌱 por <a href="https://instagram.com/chrisporto80" target="_blank">@chrisporto80</a>
  &nbsp;|&nbsp; VegSP — Guia vegano e vegetariano de São Paulo
</div>
""", unsafe_allow_html=True)

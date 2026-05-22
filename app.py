import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import date
import warnings, os, re

warnings.filterwarnings("ignore")

# ─── CONFIG ────────────────────────────────────────────────────────────────────
GSHEET_ID  = "1S89NCrisRrY9lXt1Jsko8IjQzTLFFo14C918qZvFJHM"
SHEET_NAME = "Base"
CREDS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                          "..", "PROJETO CREDITO-SINISTROS",
                          "avla-sinistros-b79c0dbac77e.json")

AZUL      = "#1565C0"
AZUL_MED  = "#1976D2"
AZUL_3    = "#42A5F5"
AZUL_LIGHT= "#EEF4FF"
CINZA_BG  = "#F5F6FA"
BRANCO    = "#FFFFFF"
BORDA     = "#DDE3EA"
TEXTO     = "#37474F"
PRETO     = "#212121"
VERDE     = "#1B8A4C"
LARANJA   = "#D4720A"
VERMELHO  = "#C0392B"
TEAL      = "#00838F"
ROXO      = "#6A1B9A"

# Gradiente azul por ranking (mais escuro = maior)
RANK_COLORS = ["#003087","#1565C0","#1976D2","#1E88E5","#42A5F5",
               "#64B5F6","#90CAF9","#BBDEFB","#E3F2FD","#EEF4FF"]
PALETTE     = ["#003087","#1565C0","#1976D2","#1E88E5","#42A5F5",
               "#00838F","#00ACC1","#1B8A4C","#43A047","#D4720A",
               "#FB8C00","#C0392B","#6A1B9A"]

# ─── PAGE ──────────────────────────────────────────────────────────────────────
_logo = os.path.join(os.path.dirname(os.path.abspath(__file__)), "avla_logo.png")
_icon = _logo if os.path.exists(_logo) else "📊"

st.set_page_config(page_title="Sinistros Crédito · AVLA",
                   page_icon=_icon, layout="wide",
                   initial_sidebar_state="collapsed")

# ─── FONTES CUSTOM (base64 em runtime) ────────────────────────────────────────
import base64 as _b64f

def _load_font_face(name, path, fmt="truetype"):
    if not os.path.exists(path): return ""
    with open(path, "rb") as _ff:
        b64 = _b64f.b64encode(_ff.read()).decode()
    return (f"@font-face {{ font-family:'{name}';"
            f" src:url('data:font/{fmt};base64,{b64}') format('{fmt}');"
            f" font-weight:normal; font-style:normal; }}")

_hal_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "HALTimezone-Regular.ttf")
_sty_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "StyreneALC-Medium.otf")

_font_face_css = (
    _load_font_face("HALTimezone", _hal_path, "truetype") + "\n  " +
    _load_font_face("StyreneALC",  _sty_path, "opentype")
)

# ─── HEADER IMAGE (base64) ────────────────────────────────────────────────────
_hdr_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "header_sinistros.png")
_hdr_img  = ""
if os.path.exists(_hdr_path):
    with open(_hdr_path, "rb") as _ff:
        _hdr_img = (f'<img src="data:image/png;base64,{_b64f.b64encode(_ff.read()).decode()}"'
                    f' style="height:44px;vertical-align:middle">')

# ─── CSS ───────────────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
  {_font_face_css}

  [data-testid="stAppViewContainer"] {{ background:{CINZA_BG}; }}
  [data-testid="stHeader"] {{ background:transparent; }}
  .block-container {{ padding:0 2rem 1rem 2rem; max-width:100%; }}
  div[data-testid="column"] {{ padding:0 0.3rem; }}

  .avla-header {{
    background:{AZUL}; padding:.55rem 2rem;
    margin:0 -2rem 1rem -2rem;
    display:flex; align-items:center; justify-content:space-between;
  }}
  .avla-header span {{ color:rgba(255,255,255,.8); font-size:.72rem;
    font-family:'StyreneALC', Arial, sans-serif; }}

  /* ── Cards KPI (mesma base, dois tamanhos) ──────────────── */
  .kpi-card {{
    background:{BRANCO}; border:1px solid {BORDA}; border-radius:8px;
    padding:.75rem 1rem; height:100%;
    font-family:'HALTimezone', Arial, sans-serif;
  }}
  .kpi-card-sm {{
    background:{BRANCO}; border:1px solid {BORDA}; border-radius:8px;
    padding:.5rem .9rem; height:100%;
    font-family:'HALTimezone', Arial, sans-serif;
    border-top-width:3px;
  }}
  .kpi-label {{
    font-size:.67rem; font-weight:700; color:{TEXTO};
    text-transform:uppercase; letter-spacing:.07em;
    font-family:'HALTimezone', Arial, sans-serif;
    margin-bottom:.18rem;
  }}
  .kpi-val {{
    font-size:1.45rem; font-weight:700; line-height:1.05; margin-bottom:.1rem;
    font-family:'HALTimezone', Arial, sans-serif;
  }}
  .kpi-val-sm {{
    font-size:1.05rem; font-weight:700; line-height:1.1; margin-bottom:.08rem;
    font-family:'HALTimezone', Arial, sans-serif;
  }}
  .kpi-sub {{ font-size:.67rem; color:#90A4AE;
    font-family:'StyreneALC', Arial, sans-serif; }}
  .c-blue   {{ color:{AZUL}; }}
  .c-green  {{ color:{VERDE}; }}
  .c-orange {{ color:{LARANJA}; }}
  .c-red    {{ color:{VERMELHO}; }}
  .c-teal   {{ color:{TEAL}; }}
  .c-purple {{ color:{ROXO}; }}

  .section-title {{
    font-size:.75rem; font-weight:700; text-transform:uppercase;
    letter-spacing:.1em; color:{TEXTO};
    font-family:'HALTimezone', Arial, sans-serif;
    margin:.6rem 0 .4rem 0; padding-bottom:.3rem;
    border-bottom:2px solid {AZUL}; display:inline-block;
  }}

  /* ── Botão Atualizar ── */
  .stButton > button {{
    background:{AZUL} !important;
    color:white !important;
    border:none !important;
    border-radius:6px !important;
    font-weight:700 !important;
    font-family:'HALTimezone', Arial, sans-serif !important;
    letter-spacing:.06em !important;
    padding:.4rem 1rem !important;
    transition:background .18s ease !important;
    box-shadow:0 2px 6px rgba(21,101,192,.25) !important;
  }}
  .stButton > button:hover {{
    background:{AZUL_MED} !important;
    box-shadow:0 3px 10px rgba(21,101,192,.35) !important;
  }}
  .stButton > button:active {{
    background:#0D47A1 !important;
    box-shadow:none !important;
  }}

  .insight, .insight-w, .insight-r {{
    border-radius:0 5px 5px 0; padding:.35rem .7rem; margin:.3rem 0;
    font-size:.78rem; font-family:'StyreneALC', Arial, sans-serif; color:{TEXTO};
  }}
  .insight   {{ background:{AZUL_LIGHT}; border-left:3px solid {AZUL}; }}
  .insight b {{ color:{AZUL}; }}
  .insight-w {{ background:#FFF3E0;  border-left:3px solid {LARANJA}; }}
  .insight-w b {{ color:{LARANJA}; }}
  .insight-r {{ background:#FFEBEE;  border-left:3px solid {VERMELHO}; }}
  .insight-r b {{ color:{VERMELHO}; }}

  /* Filtros */
  [data-baseweb="select"]>div {{ background:white!important;
    border-color:{BORDA}!important; border-radius:6px!important; }}
  [data-baseweb="select"]>div:focus-within {{ border-color:{AZUL}!important;
    box-shadow:0 0 0 2px rgba(0,48,135,.15)!important; }}
  [data-baseweb="select"] span,[data-baseweb="select"]>div>div {{
    color:{TEXTO}!important; font-family:'HALTimezone',Arial,sans-serif!important;
    font-size:.83rem!important; }}
  [data-baseweb="tag"] {{ background:{AZUL}!important; border-radius:4px!important;
    border:none!important; }}
  [data-baseweb="tag"] span {{ color:white!important; font-size:.76rem!important; }}
  [data-baseweb="popover"] ul {{ background:white!important;
    border:1px solid {BORDA}!important; border-radius:6px!important;
    box-shadow:0 4px 16px rgba(0,0,0,.1)!important; }}
  [data-baseweb="popover"] [role="option"] {{ background:white!important;
    color:{TEXTO}!important; font-family:'HALTimezone',Arial,sans-serif!important;
    font-size:.83rem!important; }}
  [data-baseweb="popover"] [role="option"]:hover {{ background:{AZUL_LIGHT}!important; }}
  [data-baseweb="popover"] [aria-selected="true"] {{
    background:#DDEAFF!important; color:{AZUL}!important; font-weight:600!important; }}
  .stSelectbox label,.stMultiSelect label {{ color:{TEXTO}!important;
    font-weight:700!important; font-size:.68rem!important;
    text-transform:uppercase!important; letter-spacing:.05em!important;
    font-family:'HALTimezone',Arial,sans-serif!important; }}

  /* Tabelas */
  div[data-testid="stDataFrame"] {{ border:1px solid {BORDA};
    border-radius:8px; overflow:hidden; }}
  div[data-testid="stDataFrame"] th,
  div[data-testid="stDataFrame"] [role="columnheader"] {{
    background:{AZUL}!important; color:white!important;
    font-weight:700!important; font-size:10px!important;
    text-transform:uppercase!important; letter-spacing:.06em!important;
    font-family:'HALTimezone',Arial,sans-serif!important; }}

  /* Tabs */
  .stTabs [data-baseweb="tab-list"] {{
    gap:0; border-bottom:2px solid {BORDA}; margin-bottom:.5rem;
  }}
  .stTabs [data-baseweb="tab"] {{
    background:transparent; border:none; border-radius:0; padding:.4rem .9rem;
    font-family:'HALTimezone',Arial,sans-serif; font-size:.78rem; font-weight:600;
    color:#90A4AE; letter-spacing:.06em; text-transform:uppercase;
  }}
  .stTabs [aria-selected="true"] {{
    color:{AZUL}!important; border-bottom:2px solid {AZUL}!important;
  }}
  .stTabs [data-baseweb="tab-panel"] {{ padding:.2rem 0 0 0; }}

  .fonte-tag {{ font-size:.62rem; color:#90A4AE;
    font-family:'StyreneALC', Arial, sans-serif;
    text-align:right; margin-top:.2rem; }}
</style>
""", unsafe_allow_html=True)


# ─── DATA ──────────────────────────────────────────────────────────────────────
@st.cache_data(ttl=300, show_spinner=False)
def load_data() -> pd.DataFrame:
    try:
        import gspread, json
        from google.oauth2.service_account import Credentials
        scopes = ["https://spreadsheets.google.com/feeds",
                  "https://www.googleapis.com/auth/drive"]
        # Streamlit Cloud: credenciais via st.secrets
        if "gsheet_credentials" in st.secrets:
            creds_dict = json.loads(st.secrets["gsheet_credentials"])
            creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        else:
            # Local: arquivo JSON
            creds = Credentials.from_service_account_file(
                os.path.normpath(CREDS_FILE), scopes=scopes)
        aba = gspread.authorize(creds).open_by_key(GSHEET_ID).worksheet(SHEET_NAME)
        all_v = aba.get_all_values()
        if len(all_v) < 2:
            return pd.DataFrame()
        headers = all_v[1]
        records = [dict(zip(headers, r + [""]*(len(headers)-len(r)))) for r in all_v[2:]]
        return pd.DataFrame(records)
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return pd.DataFrame()


def clean_valor(v):
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[R$\s]", "", str(v).strip())
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except:
        return 0.0


def fmt_brl(v):
    if v >= 1_000_000: return f"R$ {v/1_000_000:.1f}M"
    if v >= 1_000:     return f"R$ {v/1_000:.1f}K"
    return f"R$ {v:,.0f}".replace(",", ".")


def fmt_brl_full(v):
    s = f"{v:,.2f}".replace(",","X").replace(".",",").replace("X",".")
    return f"R$ {s}"


def parse_mes_ano(v):
    v = str(v).strip()
    m = re.match(r"^(\d{1,2})[/\-](\d{4})$", v)
    if m: return f"{int(m.group(1)):02d}/{m.group(2)}"
    m = re.match(r"^(\d{4})[/\-](\d{1,2})$", v)
    if m: return f"{int(m.group(2)):02d}/{m.group(1)}"
    return v


def to_date(v):
    try:
        parts = v.split("/")
        return date(int(parts[1]), int(parts[0]), 1)
    except:
        return date(2000, 1, 1)


def styled_table(df, max_h=280):
    """Tabela HTML com cabeçalho AVLA #003087 e linhas brancas."""
    TH = ("background:#1565C0;color:white;padding:6px 10px;text-align:left;"
          "font-weight:700;font-size:10px;text-transform:uppercase;"
          "letter-spacing:.06em;font-family:'HALTimezone',Arial,sans-serif;white-space:nowrap;"
          "border-right:1px solid #1976D2;")
    TD = ("padding:5px 10px;color:#212121;background:white;"
          "border-bottom:1px solid #EEF0F5;font-family:'StyreneALC',Arial,sans-serif;"
          "font-size:12px;font-weight:400;")
    header = "".join(f'<th style="{TH}">{c}</th>' for c in df.columns)
    rows   = "".join(
        "<tr>" + "".join(
            f'<td style="{TD}background:{"#F5F6FA" if i%2==0 else "white"};">{v}</td>'
            for v in row
        ) + "</tr>"
        for i, row in enumerate(df.values)
    )
    return (f'<div style="overflow-y:auto;max-height:{max_h}px;'
            f'border:1px solid #DDE3EA;border-radius:8px;">'
            f'<table style="width:100%;border-collapse:collapse;">'
            f'<thead><tr>{header}</tr></thead>'
            f'<tbody>{rows}</tbody></table></div>')


# ─── PREPARO ───────────────────────────────────────────────────────────────────
with st.spinner("Carregando dados..."):
    df_raw = load_data()

if df_raw.empty:
    st.warning("Sem dados. Verifique credenciais.")
    st.stop()

df = df_raw.copy()

COL_MAP = {
    "Setor":          ["SETOR","Setor","setor"],
    "Subsetor":       ["SUBSETOR","Subsetor","subsetor"],
    "Segurado":       ["Segurado","segurado"],
    "Apólice":        ["Apólice","Apolice","apolice"],
    "Devedor":        ["Devedor","devedor"],
    "Grupo":          ["Grupo Econômico","Grupo Economico","grupo"],
    "Mes_Ano":        ["Mes_Ano","Mês/Ano","Mês_Ano","mes_ano"],
    "Moeda":          ["Moeda","moeda"],
    "Valor_BRL":      ["Valor BRL","valor_brl","ValorBRL"],
    "Discricionario": ["Discricionário","Discricionario","discricionario"],
    "Origem":         ["Origem","origem"],
}

def find_col(df, cands):
    for c in cands:
        if c in df.columns: return c
    return None

rename = {}
for key, cands in COL_MAP.items():
    c = find_col(df, cands)
    if c and c != key:
        rename[c] = key

df = df.rename(columns=rename)
for key in COL_MAP:
    if key not in df.columns:
        df[key] = ""

df["Valor_BRL"] = df["Valor_BRL"].apply(clean_valor)
df = df[df["Valor_BRL"] > 0].copy()

df["Discricionario"] = df["Discricionario"].astype(str).str.strip().str.upper()
df["Discricionario"] = df["Discricionario"].map(
    lambda x: "Sim" if x in ("SIM","TRUE","VERDADEIRO","1","S","X") else "Não")

df["Mes_Ano"] = df["Mes_Ano"].apply(parse_mes_ano)
df["_data"]   = df["Mes_Ano"].apply(to_date)

# ─── HEADER ────────────────────────────────────────────────────────────────────
hoje = date.today()

# Carrega imagem de título em base64
import base64 as _b64
_titulo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "titulo_sinistros.png")
_titulo_html = ""
if os.path.exists(_titulo_path):
    with open(_titulo_path, "rb") as _f:
        _titulo_html = (
            f'<img src="data:image/png;base64,{_b64.b64encode(_f.read()).decode()}"'
            f' style="height:38px;vertical-align:middle">'
        )
else:
    _titulo_html = '<span style="color:white;font-size:1.2rem;font-weight:700;font-family:Arial">Sinistros de Crédito · 2026</span>'

st.markdown(f"""
<div class="avla-header">
  <div>{_hdr_img}</div>
  <span>Atualizado {hoje.strftime('%d/%m/%Y')} &nbsp;|&nbsp; Google Sheets</span>
</div>""", unsafe_allow_html=True)

# ─── FILTROS ───────────────────────────────────────────────────────────────────
fc1, fc2, fc3, fc4, fc5 = st.columns([2,2,2,2,1])

with fc1:
    periodo = st.selectbox("Período",
        ["YTD (Ano atual)","Mês atual","Últimos 3 meses","Mês específico","Todos"], index=0)
with fc2:
    setores_sel = st.multiselect("Setor",
        sorted(df["Setor"].dropna().unique()), placeholder="Todos")
with fc3:
    grupos_sel = st.multiselect("Grupo Econômico",
        sorted(df["Grupo"].dropna().unique()), placeholder="Todos")
with fc4:
    disc_sel = st.selectbox("Discricionário", ["Todos","Sim","Não"], index=0)
with fc5:
    if st.button("🔄 Atualizar", use_container_width=True):
        st.cache_data.clear(); st.rerun()

# Seletor de mês(es) — só aparece quando "Mês específico" está selecionado
meses_disp_ord = sorted(df["Mes_Ano"].unique(), key=to_date)
meses_sel = []
if periodo == "Mês específico":
    meses_sel = st.multiselect(
        "Selecionar mês(es)",
        options=meses_disp_ord,
        default=[meses_disp_ord[-1]] if meses_disp_ord else [],
        placeholder="Escolha um ou mais meses…",
    )

# ─── FILTROS: APLICA ───────────────────────────────────────────────────────────
dff = df.copy()
ano_a, mes_a = hoje.year, hoje.month

if periodo == "YTD (Ano atual)":
    dff = dff[dff["_data"].apply(lambda d: d.year == ano_a)]
elif periodo == "Mês atual":
    dff = dff[dff["_data"].apply(lambda d: d.year==ano_a and d.month==mes_a)]
elif periodo == "Últimos 3 meses":
    m3 = mes_a - 2; y3 = ano_a
    if m3 <= 0: m3 += 12; y3 -= 1
    dff = dff[dff["_data"] >= date(y3, m3, 1)]
elif periodo == "Mês específico":
    if meses_sel:
        dff = dff[dff["Mes_Ano"].isin(meses_sel)]

if setores_sel: dff = dff[dff["Setor"].isin(setores_sel)]
if grupos_sel:  dff = dff[dff["Grupo"].isin(grupos_sel)]
if disc_sel != "Todos": dff = dff[dff["Discricionario"] == disc_sel]

if dff.empty:
    st.warning("Sem dados para o filtro selecionado.")
    st.stop()

# ─── MÉTRICAS ─────────────────────────────────────────────────────────────────
total_sin    = len(dff)
valor_total  = dff["Valor_BRL"].sum()
ticket_medio = dff["Valor_BRL"].mean()
maior_sin    = dff["Valor_BRL"].max()
n_setores    = dff["Setor"].nunique()
n_grupos     = dff["Grupo"].nunique()
n_segurados  = dff["Segurado"].nunique()
disc_mask    = dff["Discricionario"] == "Sim"
disc_val     = dff[disc_mask]["Valor_BRL"].sum()
disc_pct     = disc_val / valor_total * 100 if valor_total else 0
disc_qtd_pct = disc_mask.sum() / total_sin * 100 if total_sin else 0

# ─── KPI ROW 1 ─────────────────────────────────────────────────────────────────
def card(label, val, sub, cls):
    return f"""<div class="kpi-card">
  <div class="kpi-label">{label}</div>
  <div class="kpi-val {cls}">{val}</div>
  <div class="kpi-sub">{sub}</div>
</div>"""

def mini(label, val, sub, cls, accent="#003087"):
    return f"""<div class="kpi-card-sm" style="border-top:3px solid {accent}">
  <div class="kpi-label">{label}</div>
  <div class="kpi-val-sm {cls}">{val}</div>
  <div class="kpi-sub">{sub}</div>
</div>"""

k1,k2,k3,k4 = st.columns(4)
with k1: st.markdown(card("Total Sinistros", f"{total_sin:,}", f"{n_segurados} segurados únicos", "c-blue"), unsafe_allow_html=True)
with k2: st.markdown(card("Valor Total BRL", fmt_brl(valor_total), fmt_brl_full(valor_total), "c-green"), unsafe_allow_html=True)
with k3: st.markdown(card("Ticket Médio", fmt_brl(ticket_medio), "média por sinistro", "c-orange"), unsafe_allow_html=True)
with k4: st.markdown(card("Maior Sinistro", fmt_brl(maior_sin), fmt_brl_full(maior_sin), "c-red"), unsafe_allow_html=True)

st.markdown("<div style='margin-top:.35rem'></div>", unsafe_allow_html=True)

m1,m2,m3,m4 = st.columns(4)
with m1: st.markdown(mini("Setores Afetados", str(n_setores), f"{n_grupos} grupos econômicos", "c-blue", AZUL), unsafe_allow_html=True)
with m2: st.markdown(mini("Segurados Únicos", str(n_segurados), "empresas distintas", "c-teal", TEAL), unsafe_allow_html=True)
with m3: st.markdown(mini("Discricionário — % Valor", f"{disc_pct:.1f}%", fmt_brl(disc_val), "c-orange", LARANJA), unsafe_allow_html=True)
with m4: st.markdown(mini("Discricionário — % Qtd", f"{disc_qtd_pct:.1f}%", f"{disc_mask.sum()} de {total_sin} sinistros", "c-purple", ROXO), unsafe_allow_html=True)

st.markdown("<div style='margin-top:.5rem'></div>", unsafe_allow_html=True)

# ─── PRÉ-AGREGA ────────────────────────────────────────────────────────────────
# Mensal
agg_mes = (dff.groupby("Mes_Ano")
           .agg(Valor=("Valor_BRL","sum"), Qtd=("Valor_BRL","count"))
           .reset_index().assign(_ord=lambda x: x["Mes_Ano"].apply(to_date))
           .sort_values("_ord")
           .assign(Ticket=lambda x: x["Valor"] / x["Qtd"]))

# Sazonalidade — base completa (df), agrupada por mês do ano
MESES_PT = {1:"Jan",2:"Fev",3:"Mar",4:"Abr",5:"Mai",6:"Jun",
            7:"Jul",8:"Ago",9:"Set",10:"Out",11:"Nov",12:"Dez"}
df["_mes_num"] = df["_data"].apply(lambda d: d.month)
agg_saz = (df.groupby("_mes_num")
           .agg(Valor=("Valor_BRL","sum"), Qtd=("Valor_BRL","count"))
           .reset_index().sort_values("_mes_num")
           .assign(Mes_Nome=lambda x: x["_mes_num"].map(MESES_PT),
                   Ticket=lambda x: x["Valor"] / x["Qtd"]))

# Setor
agg_set = (dff.groupby("Setor")
           .agg(Qtd=("Valor_BRL","count"), Valor=("Valor_BRL","sum"),
                Ticket=("Valor_BRL","mean"), Maior=("Valor_BRL","max"))
           .reset_index().sort_values("Valor", ascending=False))
agg_set["%V"] = agg_set["Valor"] / valor_total * 100
agg_set["%Q"] = agg_set["Qtd"]   / total_sin  * 100

# Subsetor
agg_sub = (dff[dff["Subsetor"].str.strip()!=""]
           .groupby(["Setor","Subsetor"])
           .agg(Qtd=("Valor_BRL","count"), Valor=("Valor_BRL","sum"))
           .reset_index().sort_values("Valor", ascending=False))
if not agg_sub.empty:
    agg_sub["%V"] = agg_sub["Valor"] / valor_total * 100

# Segurados
agg_seg = (dff.groupby("Segurado")
           .agg(Qtd=("Valor_BRL","count"), Valor=("Valor_BRL","sum"))
           .reset_index().sort_values("Valor", ascending=False))
agg_seg["%V"] = agg_seg["Valor"] / valor_total * 100
agg_seg["%Ac"] = agg_seg["%V"].cumsum()

# Grupos
agg_grp = (dff[dff["Grupo"].str.strip()!=""]
           .groupby("Grupo")
           .agg(Qtd=("Valor_BRL","count"), Valor=("Valor_BRL","sum"))
           .reset_index().sort_values("Valor", ascending=False))
if not agg_grp.empty:
    agg_grp["%V"]  = agg_grp["Valor"] / valor_total * 100
    agg_grp["%Ac"] = agg_grp["%V"].cumsum()

# Devedores
agg_dev = (dff[dff["Devedor"].str.strip()!=""]
           .groupby("Devedor")
           .agg(Qtd=("Valor_BRL","count"), Valor=("Valor_BRL","sum"))
           .reset_index().sort_values("Valor", ascending=False))
if not agg_dev.empty:
    agg_dev["%V"] = agg_dev["Valor"] / valor_total * 100

# Origem
agg_ori = (dff[dff["Origem"].str.strip()!=""]
           .groupby("Origem")
           .agg(Qtd=("Valor_BRL","count"), Valor=("Valor_BRL","sum"))
           .reset_index().sort_values("Valor", ascending=False))
if not agg_ori.empty:
    agg_ori["%V"] = agg_ori["Valor"] / valor_total * 100

# Top maiores casos
top_casos = (dff.nlargest(10,"Valor_BRL")
             [["Mes_Ano","Segurado","Devedor","Setor","Apólice","Valor_BRL"]]
             .copy())
top_casos["%T"] = top_casos["Valor_BRL"] / valor_total * 100

# Discricionário
disc_agg = (dff.groupby("Discricionario")
            .agg(Valor=("Valor_BRL","sum"), Qtd=("Valor_BRL","count"))
            .reset_index())

# Faixas
bins   = [0,10_000,50_000,100_000,500_000,1_000_000,float("inf")]
labels = ["< 10K","10K–50K","50K–100K","100K–500K","500K–1M","> 1M"]
dff["Faixa"] = pd.cut(dff["Valor_BRL"], bins=bins, labels=labels, right=False)
agg_fx = (dff.groupby("Faixa", observed=True)
          .agg(Qtd=("Valor_BRL","count"), Valor=("Valor_BRL","sum"))
          .reset_index())
agg_fx["%Q"] = agg_fx["Qtd"]   / total_sin  * 100
agg_fx["%V"] = agg_fx["Valor"] / valor_total * 100

# Heatmap pivot
pivot = dff.pivot_table(index="Setor", columns="Mes_Ano",
                        values="Valor_BRL", aggfunc="sum", fill_value=0)
meses_ord = sorted(pivot.columns.tolist(), key=to_date)
pivot = pivot[meses_ord].loc[pivot.sum(axis=1).sort_values(ascending=False).index]


# ─── HELPERS DE LAYOUT ─────────────────────────────────────────────────────────
FIG_FONT = dict(family="'HALTimezone', Arial, sans-serif", color=PRETO)

def base_layout(height=285, t=8, b=8, l=8, r=8):
    return dict(height=height, margin=dict(t=t,b=b,l=l,r=r),
                plot_bgcolor=BRANCO, paper_bgcolor=BRANCO,
                font=FIG_FONT)

def x_axis(angle=0, fmt=None):
    d = dict(showgrid=False, tickfont=dict(family="'HALTimezone', Arial",color=PRETO,size=10),
             tickangle=angle)
    if fmt: d["tickformat"] = fmt
    return d

def y_axis(grid=True, fmt=None, suffix=None):
    d = dict(showgrid=grid, gridcolor="#EEF0F5",
             tickfont=dict(family="'HALTimezone', Arial",color=PRETO,size=10))
    if fmt:    d["tickformat"] = fmt
    if suffix: d["ticksuffix"] = suffix
    return d

def rank_color(i, n=None):
    return RANK_COLORS[min(i, len(RANK_COLORS)-1)]


# ═══════════════════════════════════════════════════════════════════════════════
# TABS
# ═══════════════════════════════════════════════════════════════════════════════
tab1, tab2, tab3, tab4 = st.tabs([
    "📊  Visão Geral",
    "🏭  Por Setor",
    "🎯  Concentração",
    "📋  Análise",
])


# ──────────────────────────────────────────────────────────────────────────────
# TAB 1 — VISÃO GERAL
# ──────────────────────────────────────────────────────────────────────────────
with tab1:
    c1, c2 = st.columns([3, 2])

    with c1:
        st.markdown("<div class='section-title'>Evolução Mensal — Valor & Nº Sinistros</div>",
                    unsafe_allow_html=True)
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_bar(
            x=agg_mes["Mes_Ano"], y=agg_mes["Valor"],
            marker_color=AZUL, name="Valor BRL",
            text=agg_mes["Valor"].apply(fmt_brl), textposition="outside",
            textfont=dict(family="'HALTimezone', Arial", color=PRETO, size=10),
            hovertemplate="<b>%{x}</b><br>Valor: %{text}<extra></extra>",
            secondary_y=False,
        )
        fig.add_scatter(
            x=agg_mes["Mes_Ano"], y=agg_mes["Qtd"],
            mode="lines+markers+text", name="Nº Sinistros",
            line=dict(color=LARANJA, width=2.5),
            marker=dict(size=7, color=LARANJA),
            text=agg_mes["Qtd"], textposition="top center",
            textfont=dict(family="'HALTimezone', Arial", color=LARANJA, size=10, weight="bold"),
            hovertemplate="<b>%{x}</b><br>Qtd: %{y}<extra></extra>",
            secondary_y=True,
        )
        fig.update_layout(**base_layout(290, t=5, r=15),
                          showlegend=True,
                          legend=dict(orientation="h",y=1.08,x=1,xanchor="right",
                                      font=dict(family="'HALTimezone', Arial",color=PRETO,size=10)))
        fig.update_xaxes(showgrid=False,
                         tickfont=dict(family="'HALTimezone', Arial",color=PRETO,size=10))
        fig.update_yaxes(showgrid=True, gridcolor="#EEF0F5",
                         tickfont=dict(family="'HALTimezone', Arial",color=PRETO,size=10),
                         tickformat=",.0f", secondary_y=False)
        fig.update_yaxes(showgrid=False,
                         tickfont=dict(family="'HALTimezone', Arial",color=PRETO,size=10),
                         secondary_y=True)
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})

    with c2:
        st.markdown("<div class='section-title'>🏆 Top 10 — Maiores Sinistros</div>",
                    unsafe_allow_html=True)
        tc = top_casos.copy()
        tc.insert(0, "#", range(1, len(tc)+1))
        tc_show = tc[["#","Mes_Ano","Segurado","Setor","Valor_BRL","%T"]].copy()
        tc_show.columns = ["#","Mês","Segurado","Setor","Valor BRL","% Total"]
        tc_show["Valor BRL"] = tc_show["Valor BRL"].apply(fmt_brl_full)
        tc_show["% Total"]   = tc_show["% Total"].apply(lambda x: f"{x:.1f}%")
        st.markdown(styled_table(tc_show, max_h=295), unsafe_allow_html=True)
        st.markdown(f"<div class='fonte-tag'>Top 10 de {total_sin:,} registros</div>",
                    unsafe_allow_html=True)

    # ── Linha 2: Ticket Médio + Sazonalidade ──────────────────────────────────
    st.markdown("<div style='margin-top:.2rem'></div>", unsafe_allow_html=True)
    t1, t2 = st.columns([3, 2])

    with t1:
        st.markdown("<div class='section-title'>Evolução do Ticket Médio por Mês</div>",
                    unsafe_allow_html=True)
        fig_tk = go.Figure()
        fig_tk.add_scatter(
            x=agg_mes["Mes_Ano"], y=agg_mes["Ticket"],
            mode="lines+markers+text",
            line=dict(color=TEAL, width=2.5),
            marker=dict(size=8, color=TEAL),
            text=agg_mes["Ticket"].apply(fmt_brl),
            textposition="top center",
            textfont=dict(family="'HALTimezone', Arial", color=PRETO, size=10),
            hovertemplate="<b>%{x}</b><br>Ticket: %{text}<extra></extra>",
            fill="tozeroy", fillcolor="rgba(0,131,143,0.07)",
            name="Ticket Médio",
        )
        # Linha de média geral
        media_geral = ticket_medio
        fig_tk.add_hline(y=media_geral, line_dash="dash",
                         line_color=LARANJA, line_width=1.5, opacity=0.8,
                         annotation_text=f"Média: {fmt_brl(media_geral)}",
                         annotation_position="right",
                         annotation_font=dict(color=LARANJA, size=9, family="'HALTimezone', Arial"))
        fig_tk.update_layout(
            height=215, margin=dict(t=5, b=8, l=8, r=100),
            plot_bgcolor=BRANCO, paper_bgcolor=BRANCO,
            font=FIG_FONT, showlegend=False,
            xaxis=dict(showgrid=False,
                       tickfont=dict(family="'HALTimezone', Arial", color=PRETO, size=10)),
            yaxis=dict(showgrid=True, gridcolor="#EEF0F5",
                       tickfont=dict(family="'HALTimezone', Arial", color=PRETO, size=10),
                       tickformat=",.0f"),
        )
        st.plotly_chart(fig_tk, use_container_width=True,
                        config={"displayModeBar": False})

    with t2:
        st.markdown("<div class='section-title'>Sazonalidade — Base Completa</div>",
                    unsafe_allow_html=True)
        # Intensidade de cor por valor (mais escuro = mês mais pesado)
        max_v = agg_saz["Valor"].max() if not agg_saz.empty else 1
        saz_colors = [
            f"rgba(0,48,135,{0.25 + 0.75*(v/max_v):.2f})"
            for v in agg_saz["Valor"]
        ]
        fig_saz = go.Figure(go.Bar(
            x=agg_saz["Mes_Nome"], y=agg_saz["Qtd"],
            marker_color=saz_colors,
            text=agg_saz["Qtd"], textposition="outside",
            textfont=dict(family="'HALTimezone', Arial", color=PRETO, size=10),
            customdata=agg_saz["Valor"].apply(fmt_brl),
            hovertemplate=(
                "<b>%{x}</b><br>Sinistros: %{y}<br>Valor: %{customdata}<extra></extra>"
            ),
            name="Nº Sinistros",
        ))
        fig_saz.update_layout(
            height=215, margin=dict(t=5, b=8, l=8, r=8),
            plot_bgcolor=BRANCO, paper_bgcolor=BRANCO,
            font=FIG_FONT, showlegend=False,
            xaxis=dict(showgrid=False,
                       tickfont=dict(family="'HALTimezone', Arial", color=PRETO, size=10)),
            yaxis=dict(showgrid=True, gridcolor="#EEF0F5",
                       tickfont=dict(family="'HALTimezone', Arial", color=PRETO, size=10)),
            annotations=[dict(
                text="Intensidade de cor = peso em valor", x=1, y=1.08,
                xref="paper", yref="paper", showarrow=False,
                font=dict(family="'HALTimezone', Arial", color="#90A4AE", size=9),
                xanchor="right",
            )],
        )
        st.plotly_chart(fig_saz, use_container_width=True,
                        config={"displayModeBar": False})


# ──────────────────────────────────────────────────────────────────────────────
# TAB 2 — POR SETOR
# ──────────────────────────────────────────────────────────────────────────────
with tab2:
    # Linha 1: tabela + barras
    s1, s2 = st.columns([2, 3])

    with s1:
        st.markdown("<div class='section-title'>Resumo por Setor</div>",
                    unsafe_allow_html=True)
        tbl = pd.DataFrame({
            "Setor":      agg_set["Setor"],
            "Qtd":        agg_set["Qtd"],
            "% Qtd":      agg_set["%Q"].apply(lambda x: f"{x:.1f}%"),
            "Valor BRL":  agg_set["Valor"].apply(fmt_brl),
            "% Valor":    agg_set["%V"].apply(lambda x: f"{x:.1f}%"),
            "Ticket Méd": agg_set["Ticket"].apply(fmt_brl),
            "Maior":      agg_set["Maior"].apply(fmt_brl),
        })
        st.markdown(styled_table(tbl, max_h=310), unsafe_allow_html=True)

    with s2:
        st.markdown("<div class='section-title'>% Valor vs % Qtd por Setor</div>",
                    unsafe_allow_html=True)
        srt = agg_set.sort_values("%V", ascending=True)
        fig_s = go.Figure()
        fig_s.add_bar(
            y=srt["Setor"], x=srt["%V"], orientation="h", name="% Valor",
            marker_color=[rank_color(i) for i in range(len(srt)-1, -1, -1)],
            text=srt["%V"].apply(lambda x: f"{x:.1f}%"), textposition="outside",
            textfont=dict(family="'HALTimezone', Arial", color=PRETO, size=10),
            hovertemplate="<b>%{y}</b><br>% Valor: %{text}<br>Valor: %{customdata}<extra></extra>",
            customdata=srt["Valor"].apply(fmt_brl_full),
        )
        fig_s.add_bar(
            y=srt["Setor"], x=srt["%Q"], orientation="h", name="% Qtd (sinistros)",
            marker_color=LARANJA, opacity=0.55,
            hovertemplate="<b>%{y}</b><br>% Qtd: %{x:.1f}%<extra></extra>",
        )
        fig_s.update_layout(
            height=310, margin=dict(t=8, b=52, l=8, r=80),
            plot_bgcolor=BRANCO, paper_bgcolor=BRANCO, font=FIG_FONT,
            barmode="overlay", showlegend=True,
            legend=dict(orientation="h", y=-0.16, x=0.5, xanchor="center",
                        font=dict(family="'HALTimezone', Arial", color=PRETO, size=10)),
            xaxis=dict(showgrid=True, gridcolor="#EEF0F5", ticksuffix="%",
                       tickfont=dict(family="'HALTimezone', Arial", color=PRETO, size=10)),
            yaxis=dict(showgrid=False,
                       tickfont=dict(family="'HALTimezone', Arial", color=PRETO, size=10)))
        st.plotly_chart(fig_s, use_container_width=True, config={"displayModeBar":False})

    # Insight setor
    top1 = agg_set.iloc[0]
    top2pct = agg_set.head(2)["%V"].sum() if len(agg_set) >= 2 else top1["%V"]
    disc_s = dff[disc_mask].groupby("Setor")["Valor_BRL"].sum()
    disc_s_name = disc_s.idxmax() if not disc_s.empty else "–"
    ia, ib = st.columns(2)
    with ia:
        st.markdown(f"""<div class='insight'>
        📊 <b>{top1['Setor']}</b> lidera com <b>{top1['%V']:.1f}%</b> do valor total
        ({fmt_brl(top1['Valor'])}) · {int(top1['Qtd'])} sinistros.
        </div>""", unsafe_allow_html=True)
    with ib:
        st.markdown(f"""<div class='insight-w'>
        ⚠️ 2 maiores setores = <b>{top2pct:.1f}%</b> do total.
        {f"Maior disc.: <b>{disc_s_name}</b>" if not disc_s.empty else ""}
        </div>""", unsafe_allow_html=True)

    # Linha 2: heatmap setor × mês
    st.markdown("<div class='section-title'>Heatmap — Valor BRL · Setor × Mês</div>",
                unsafe_allow_html=True)
    text_m = [[fmt_brl(v) if v > 0 else "" for v in row] for row in pivot.values]
    fig_h = go.Figure(go.Heatmap(
        z=pivot.values, x=pivot.columns.tolist(), y=pivot.index.tolist(),
        text=text_m, texttemplate="%{text}",
        textfont=dict(family="'HALTimezone', Arial", size=9, color="#003087"),
        colorscale=[[0,"#FFFFFF"],[0.25,"#DDEAFF"],[0.6,"#90CAF9"],[1.0,"#42A5F5"]],
        showscale=False,
        hovertemplate="<b>%{y}</b> · <b>%{x}</b><br>%{text}<extra></extra>",
    ))
    fig_h.update_layout(**base_layout(max(180, len(pivot)*34+40), t=5),
                        xaxis=dict(tickfont=dict(family="'HALTimezone', Arial",color=PRETO,size=10)),
                        yaxis=dict(tickfont=dict(family="'HALTimezone', Arial",color=PRETO,size=10)))
    st.plotly_chart(fig_h, use_container_width=True, config={"displayModeBar":False})


# ──────────────────────────────────────────────────────────────────────────────
# TAB 3 — CONCENTRAÇÃO DE RISCO
# ──────────────────────────────────────────────────────────────────────────────
with tab3:
    p1, p2 = st.columns([3, 2])

    # Pareto segurados
    with p1:
        st.markdown("<div class='section-title'>Curva de Pareto — Top Segurados</div>",
                    unsafe_allow_html=True)
        ps = agg_seg.head(12).copy()
        fig_p = make_subplots(specs=[[{"secondary_y": True}]])
        fig_p.add_bar(
            x=ps["Segurado"], y=ps["%V"],
            marker_color=[rank_color(i) for i in range(len(ps))],
            name="% do Total",
            text=ps["%V"].apply(lambda x: f"{x:.1f}%"), textposition="outside",
            textfont=dict(family="'HALTimezone', Arial", color=PRETO, size=9),
            hovertemplate="<b>%{x}</b><br>% Total: %{text}<br>Valor: %{customdata}<extra></extra>",
            customdata=ps["Valor"].apply(fmt_brl_full),
            secondary_y=False,
        )
        fig_p.add_scatter(
            x=ps["Segurado"], y=ps["%Ac"],
            mode="lines+markers+text", name="% Acum.",
            line=dict(color=VERMELHO, width=2),
            marker=dict(size=6, color=VERMELHO),
            text=ps["%Ac"].apply(lambda x: f"{x:.0f}%"), textposition="top center",
            textfont=dict(family="'HALTimezone', Arial", color=VERMELHO, size=9),
            hovertemplate="<b>%{x}</b><br>% Acum: %{y:.1f}%<extra></extra>",
            secondary_y=True,
        )
        # Linha 80%
        fig_p.add_hline(y=80, line_dash="dash", line_color=LARANJA,
                        line_width=1.5, opacity=0.8)
        fig_p.update_layout(**base_layout(290, t=10, r=10), showlegend=True,
                            legend=dict(orientation="h",y=1.08,x=1,xanchor="right",
                                        font=dict(family="'HALTimezone', Arial",color=PRETO,size=9)))
        fig_p.update_xaxes(showgrid=False, tickangle=-35,
                           tickfont=dict(family="'HALTimezone', Arial",color=PRETO,size=9))
        fig_p.update_yaxes(ticksuffix="%",
                           tickfont=dict(family="'HALTimezone', Arial",color=PRETO,size=9),
                           showgrid=True, gridcolor="#EEF0F5", secondary_y=False)
        fig_p.update_yaxes(ticksuffix="%", range=[0, 110],
                           tickfont=dict(family="'HALTimezone', Arial",color=PRETO,size=9),
                           showgrid=False, secondary_y=True)
        st.plotly_chart(fig_p, use_container_width=True, config={"displayModeBar":False})

        n80 = int((ps["%Ac"] <= 80).sum()) + 1
        pct80 = ps.iloc[min(n80-1, len(ps)-1)]["%Ac"]
        st.markdown(f"""<div class='insight'>
        📌 <b>Top {n80} segurados</b> concentram <b>{pct80:.1f}%</b> do valor total.
        Linha laranja = limite 80%.
        </div>""", unsafe_allow_html=True)

    # Top devedores
    with p2:
        st.markdown("<div class='section-title'>Top Devedores — Valor BRL</div>",
                    unsafe_allow_html=True)
        if not agg_dev.empty:
            td = agg_dev.head(10).copy()
            fig_dev = go.Figure(go.Bar(
                y=td["Devedor"],
                x=td["Valor"],
                orientation="h",
                marker_color=[rank_color(i) for i in range(len(td))],
                text=[f"{fmt_brl(v)}  {p:.1f}%"
                      for v, p in zip(td["Valor"], td["%V"])],
                textposition="outside",
                textfont=dict(family="'HALTimezone', Arial", color=PRETO, size=9),
                customdata=list(zip(td["Valor"].apply(fmt_brl_full), td["Qtd"])),
                hovertemplate=(
                    "<b>%{y}</b><br>"
                    "Valor: %{customdata[0]}<br>"
                    "Qtd: %{customdata[1]} sinistros<extra></extra>"
                ),
            ))
            fig_dev.update_layout(
                height=295, margin=dict(t=5, b=5, l=5, r=110),
                plot_bgcolor=BRANCO, paper_bgcolor=BRANCO,
                font=FIG_FONT, showlegend=False,
                xaxis=dict(showgrid=True, gridcolor="#EEF0F5", tickformat=",.0f",
                           tickfont=dict(family="'HALTimezone', Arial", color=PRETO, size=9)),
                yaxis=dict(showgrid=False, autorange="reversed",
                           tickfont=dict(family="'HALTimezone', Arial", color=PRETO, size=9)),
            )
            st.plotly_chart(fig_dev, use_container_width=True,
                            config={"displayModeBar": False})
            top_dev = td.iloc[0]
            st.markdown(f"""<div class='insight-r'>
            🔴 <b>{top_dev['Devedor']}</b>: {fmt_brl_full(top_dev['Valor'])}
            ({top_dev['%V']:.1f}% do total) · {int(top_dev['Qtd'])} sinistros.
            </div>""", unsafe_allow_html=True)
        else:
            st.info("Coluna Devedor não preenchida.")

    # Linha 2: pareto grupos
    if not agg_grp.empty:
        st.markdown("<div class='section-title'>Curva de Pareto — Grupos Econômicos</div>",
                    unsafe_allow_html=True)
        pg = agg_grp.head(12).copy()
        fig_pg = make_subplots(specs=[[{"secondary_y": True}]])
        fig_pg.add_bar(
            x=pg["Grupo"], y=pg["%V"],
            marker_color=[rank_color(i) for i in range(len(pg))],
            name="% do Total",
            text=pg["%V"].apply(lambda x: f"{x:.1f}%"), textposition="outside",
            textfont=dict(family="'HALTimezone', Arial", color=PRETO, size=9),
            hovertemplate="<b>%{x}</b><br>% Total: %{text}<br>Valor: %{customdata}<extra></extra>",
            customdata=pg["Valor"].apply(fmt_brl_full),
            secondary_y=False,
        )
        fig_pg.add_scatter(
            x=pg["Grupo"], y=pg["%Ac"],
            mode="lines+markers+text", name="% Acum.",
            line=dict(color=VERMELHO, width=2),
            marker=dict(size=6, color=VERMELHO),
            text=pg["%Ac"].apply(lambda x: f"{x:.0f}%"), textposition="top center",
            textfont=dict(family="'HALTimezone', Arial", color=VERMELHO, size=9),
            secondary_y=True,
        )
        fig_pg.add_hline(y=80, line_dash="dash", line_color=LARANJA,
                         line_width=1.5, opacity=0.8)
        fig_pg.update_layout(**base_layout(240, t=10, r=10), showlegend=False)
        fig_pg.update_xaxes(showgrid=False, tickangle=-30,
                            tickfont=dict(family="'HALTimezone', Arial",color=PRETO,size=9))
        fig_pg.update_yaxes(ticksuffix="%",
                            tickfont=dict(family="'HALTimezone', Arial",color=PRETO,size=9),
                            showgrid=True, gridcolor="#EEF0F5", secondary_y=False)
        fig_pg.update_yaxes(ticksuffix="%", range=[0,110],
                            tickfont=dict(family="'HALTimezone', Arial",color=PRETO,size=9),
                            showgrid=False, secondary_y=True)
        st.plotly_chart(fig_pg, use_container_width=True,
                        config={"displayModeBar":False})

        n80g = int((pg["%Ac"] <= 80).sum()) + 1
        pct80g = pg.iloc[min(n80g-1, len(pg)-1)]["%Ac"]
        st.markdown(f"""<div class='insight'>
        📌 <b>Top {n80g} grupos</b> concentram <b>{pct80g:.1f}%</b> do valor total.
        </div>""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# TAB 4 — ANÁLISE
# ──────────────────────────────────────────────────────────────────────────────
with tab4:
    a1, a2 = st.columns([2, 2])

    with a1:
        st.markdown("<div class='section-title'>Frequência & Valor por Faixa</div>",
                    unsafe_allow_html=True)
        fig_fx = make_subplots(specs=[[{"secondary_y": True}]])
        fig_fx.add_bar(
            x=agg_fx["Faixa"].astype(str), y=agg_fx["Qtd"],
            marker_color=AZUL, name="Nº Sinistros",
            text=agg_fx["Qtd"], textposition="outside",
            textfont=dict(family="'HALTimezone', Arial", color=PRETO, size=10),
            secondary_y=False,
        )
        fig_fx.add_scatter(
            x=agg_fx["Faixa"].astype(str), y=agg_fx["%V"],
            mode="lines+markers+text", name="% Valor",
            line=dict(color=VERMELHO, width=2),
            marker=dict(size=7, color=VERMELHO),
            text=agg_fx["%V"].apply(lambda x: f"{x:.1f}%"),
            textposition="top center",
            textfont=dict(family="'HALTimezone', Arial", color=VERMELHO, size=9),
            secondary_y=True,
        )
        fig_fx.update_layout(**base_layout(285, t=5, r=10), showlegend=True,
                             legend=dict(orientation="h",y=1.08,x=1,xanchor="right",
                                         font=dict(family="'HALTimezone', Arial",color=PRETO,size=10)))
        fig_fx.update_xaxes(showgrid=False, tickangle=-20,
                            tickfont=dict(family="'HALTimezone', Arial",color=PRETO,size=10))
        fig_fx.update_yaxes(showgrid=True, gridcolor="#EEF0F5",
                            tickfont=dict(family="'HALTimezone', Arial",color=PRETO,size=10),
                            secondary_y=False)
        fig_fx.update_yaxes(ticksuffix="%", range=[0,105],
                            tickfont=dict(family="'HALTimezone', Arial",color=PRETO,size=10),
                            showgrid=False, secondary_y=True)
        st.plotly_chart(fig_fx, use_container_width=True,
                        config={"displayModeBar":False})

        big = agg_fx[agg_fx["Faixa"].astype(str).isin(["100K–500K","500K–1M","> 1M"])]
        st.markdown(f"""<div class='insight-r'>
        🔴 Sinistros <b>&gt; R$ 100K</b>: {int(big['Qtd'].sum())} casos ({big['%Q'].sum():.1f}%
        da freq.) mas <b>{big['%V'].sum():.1f}% do valor total</b>.
        </div>""", unsafe_allow_html=True)

    with a2:
        st.markdown("<div class='section-title'>Discricionário — Comparativo</div>",
                    unsafe_allow_html=True)

        # Stacked 100% horizontal: % Valor e % Qtd, Sim vs Não
        disc_sim = dff[disc_mask]
        disc_nao = dff[~disc_mask]
        t_sim_v  = disc_sim["Valor_BRL"].sum()
        t_nao_v  = disc_nao["Valor_BRL"].sum()
        t_sim_q  = len(disc_sim)
        t_nao_q  = len(disc_nao)
        t_sim_tk = disc_sim["Valor_BRL"].mean() if t_sim_q > 0 else 0
        t_nao_tk = disc_nao["Valor_BRL"].mean() if t_nao_q > 0 else 0

        metricas = ["% do Valor Total", "% dos Sinistros"]
        nao_pcts = [t_nao_v / valor_total * 100, t_nao_q / total_sin * 100]
        sim_pcts = [t_sim_v / valor_total * 100, t_sim_q / total_sin * 100]

        fig_d = go.Figure()
        fig_d.add_bar(
            y=metricas, x=nao_pcts, orientation="h",
            name="Não-Disc.", marker_color=AZUL,
            text=[f"{v:.1f}%" for v in nao_pcts],
            textposition="inside",
            textfont=dict(family="'HALTimezone', Arial", color="white", size=12, weight="bold"),
            hovertemplate="<b>%{y}</b><br>Não-Disc.: %{x:.1f}%<extra></extra>",
        )
        fig_d.add_bar(
            y=metricas, x=sim_pcts, orientation="h",
            name="Discricionário", marker_color=LARANJA,
            text=[f"{v:.1f}%" for v in sim_pcts],
            textposition="inside",
            textfont=dict(family="'HALTimezone', Arial", color="white", size=12, weight="bold"),
            hovertemplate="<b>%{y}</b><br>Disc.: %{x:.1f}%<extra></extra>",
        )
        fig_d.update_layout(
            height=140, margin=dict(t=5, b=5, l=5, r=5),
            plot_bgcolor=BRANCO, paper_bgcolor=BRANCO, font=FIG_FONT,
            barmode="stack", showlegend=True,
            legend=dict(orientation="h", y=1.18, x=0.5, xanchor="center",
                        font=dict(family="'HALTimezone', Arial", color=PRETO, size=10)),
            xaxis=dict(range=[0, 100], ticksuffix="%", showgrid=False,
                       tickfont=dict(family="'HALTimezone', Arial", color=PRETO, size=10)),
            yaxis=dict(showgrid=False,
                       tickfont=dict(family="'HALTimezone', Arial", color=PRETO, size=10)),
        )
        st.plotly_chart(fig_d, use_container_width=True,
                        config={"displayModeBar": False})

        # Comparativo de ticket médio
        tk_diff_pct = ((t_sim_tk - t_nao_tk) / t_nao_tk * 100) if t_nao_tk > 0 else 0
        arrow = "↑" if tk_diff_pct > 0 else "↓"
        color_tk = VERMELHO if tk_diff_pct > 0 else VERDE
        st.markdown(f"""<div class='insight-w'>
        <b>Ticket Médio Disc.:</b> {fmt_brl(t_sim_tk)} vs {fmt_brl(t_nao_tk)} (não-disc.)<br>
        <span style="color:{color_tk};font-weight:700">{arrow} {abs(tk_diff_pct):.1f}%</span>
        {"maior" if tk_diff_pct > 0 else "menor"} que a média dos não-discricionários.
        </div>""", unsafe_allow_html=True)

        # Por Origem (se existir)
        if not agg_ori.empty and len(agg_ori) > 1:
            st.markdown("<div class='section-title'>Por Origem</div>",
                        unsafe_allow_html=True)
            ori_show = pd.DataFrame({
                "Origem":    agg_ori["Origem"],
                "Qtd":       agg_ori["Qtd"],
                "Valor BRL": agg_ori["Valor"].apply(fmt_brl),
                "% Valor":   agg_ori["%V"].apply(lambda x: f"{x:.1f}%"),
            })
            st.markdown(styled_table(ori_show, max_h=130), unsafe_allow_html=True)

    # Tabela detalhada
    with st.expander("📄 Base completa filtrada", expanded=False):
        cols_s = [c for c in ["Mes_Ano","Setor","Subsetor","Segurado","Devedor",
                               "Apólice","Grupo","Moeda","Valor_BRL",
                               "Discricionario","Origem"] if c in dff.columns]
        df_s = dff[cols_s].copy().rename(columns={
            "Mes_Ano":"Mês/Ano","Grupo":"Grupo Econômico",
            "Valor_BRL":"Valor BRL","Discricionario":"Discricionário"})
        if "Valor BRL" in df_s.columns:
            df_s["Valor BRL"] = df_s["Valor BRL"].apply(fmt_brl_full)
        st.markdown(styled_table(df_s.reset_index(drop=True), max_h=280),
                    unsafe_allow_html=True)
        st.markdown(f"<div class='fonte-tag'>{len(dff):,} registros · Fonte: Google Sheets · Sinistros Crédito 2026</div>",
                    unsafe_allow_html=True)

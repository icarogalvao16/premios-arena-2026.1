"""
========================================================================================
🏆 DASHBOARD DE CAMPEONATO DE FUTEBOL 2026 - PRÊMIOS MANGABEIRA
========================================================================================
Desenvolvido com Python, Streamlit, Pandas, Matplotlib e Seaborn.
Compatível com a planilha 'Prêmios 2026.xlsx' com as abas:
- Planilha Craque (MVP)
- Planilha Brocador (Artilharia)
- Planilha César (Melhor Goleiro)
- Planilha Reis do Sábado (Coroas de Sábado)
========================================================================================
"""

import os
import io
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import streamlit as st

# ======================================================================================
# CONFIGURAÇÃO DA PÁGINA STREAMLIT
# ======================================================================================
st.set_page_config(
    page_title="Prêmios 2026 | Dashboard de Futebol",
    page_icon="⚽",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ======================================================================================
# ESTILIZAÇÃO CSS CUSTOMIZADA (TEMA ESCURO DE ALTO CONTRASTE)
# ======================================================================================
st.markdown("""
<style>
    .stApp {
        background-color: #0b0f17;
        color: #f3f4f6;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    }
    [data-testid="stSidebar"] {
        background-color: #111827;
        border-right: 1px solid #1f2937;
    }
    .metric-card {
        background: linear-gradient(145deg, #161f30, #0f172a);
        border: 1px solid #23324a;
        border-radius: 12px;
        padding: 18px 20px;
        margin-bottom: 12px;
        box-shadow: 0 4px 14px rgba(0, 0, 0, 0.35);
        transition: transform 0.2s ease, border-color 0.2s ease;
    }
    .metric-card:hover {
        transform: translateY(-2px);
        border-color: #10b981;
    }
    .metric-rank {
        font-size: 0.75rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        margin-bottom: 4px;
    }
    .rank-gold { color: #fbbf24; }
    .rank-silver { color: #94a3b8; }
    .rank-bronze { color: #f97316; }
    .rank-crown { color: #ec4899; }
    .rank-stats { color: #38bdf8; }
    .metric-player {
        font-size: 1.35rem;
        font-weight: 800;
        color: #ffffff;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    .metric-value {
        font-size: 1.85rem;
        font-weight: 800;
        color: #10b981;
        font-variant-numeric: tabular-nums;
        line-height: 1.1;
        margin-top: 4px;
    }
    .metric-sub {
        font-size: 0.8rem;
        color: #9ca3af;
        margin-top: 4px;
    }
    [data-testid="stDataFrame"] {
        background-color: #111827;
        border-radius: 10px;
        border: 1px solid #1f2937;
    }
</style>
""", unsafe_allow_html=True)

# ======================================================================================
# CONFIGURAÇÃO DO MATPLOTLIB / SEABORN (DARK THEME)
# ======================================================================================
plt.style.use('dark_background')
plt.rcParams['font.sans-serif'] = 'DejaVu Sans'
plt.rcParams['axes.edgecolor'] = '#374151'
plt.rcParams['axes.linewidth'] = 0.8
plt.rcParams['grid.color'] = '#1f2937'
plt.rcParams['grid.linestyle'] = '--'
plt.rcParams['grid.alpha'] = 0.6

MONTHS_LIST = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]

RAW_DATA_CSV = """Jogador,Janeiro,RS 1,Fevereiro,RS 2,Março,RS 3,Abril,RS 4,Maio,RS 5,Junho,RS 6,Julho,RS 7,Agosto,RS 8,Setembro,RS 9,Outubro,RS 10,Novembro,RS 11,Dezembro,RS 12,Pontuação Geral,RS Total
Adriano,3,1,7,1,4,0,10,0,14,0,3,0,0,0,11,0,9,0,3,0,0,0,0,0,64,2
Jandir,3,1,5,0,2,0,1,0,16,0,7,0,0,0,10,0,9,0,3,0,0,0,0,0,56,1
Ballet,0,0,10,2,6,0,10,0,12,0,3,0,0,0,8,0,6,0,0,0,0,0,0,0,55,2
Rodrigo,3,1,9,1,12,0,7,0,7,0,1,0,0,0,3,0,10,0,3,0,0,0,0,0,55,2
Robério,3,1,10,1,7,0,4,0,12,0,3,0,0,0,5,0,9,0,0,0,0,0,0,0,53,2
Lucas,3,1,10,1,10,0,9,0,5,0,0,0,0,0,8,0,6,0,0,0,0,0,0,0,51,2
Talisca,0,0,3,1,10,0,4,0,14,0,6,0,0,0,7,0,7,0,0,0,0,0,0,0,51,1
Rafa,0,0,10,1,7,0,6,0,10,0,3,0,0,0,8,0,3,0,3,0,0,0,0,0,50,1
Ricardo,3,1,3,1,7,0,11,0,8,0,0,0,0,0,8,0,4,0,3,0,0,0,0,0,47,2
Canário,3,1,4,0,9,0,7,0,12,0,1,0,0,0,0,0,6,0,3,0,0,0,0,0,45,1
Remi,0,0,6,0,4,0,3,0,14,0,6,0,0,0,8,0,3,0,0,0,0,0,0,0,44,0
Albino,0,0,9,0,9,0,1,0,7,0,7,0,0,0,8,0,3,0,0,0,0,0,0,0,44,0
Raul,0,0,13,1,3,0,6,0,3,0,3,0,0,0,7,0,9,0,0,0,0,0,0,0,44,1
PJ,3,1,4,0,4,0,11,0,5,0,0,0,0,0,11,0,3,0,3,0,0,0,0,0,44,1
Minuano,0,0,7,1,11,0,6,0,11,0,0,0,0,0,0,0,6,0,3,0,0,0,0,0,44,1
Rafael Santana,0,0,3,1,9,0,3,0,10,0,4,0,0,0,5,0,9,0,0,0,0,0,0,0,43,1
Reinaldo,0,0,6,1,4,0,3,0,14,0,0,0,0,0,6,0,7,0,3,0,0,0,0,0,43,1
Amaury,0,0,7,1,7,0,1,0,6,0,3,0,0,0,5,0,13,0,0,0,0,0,0,0,42,1
Capela,0,0,7,0,0,0,6,0,12,0,0,0,0,0,6,0,7,0,3,0,0,0,0,0,41,0
M. Yuri,0,0,4,1,15,0,0,0,11,0,3,0,0,0,7,0,0,0,0,0,0,0,0,0,40,1
Vini HC,0,0,3,1,7,0,7,0,7,0,6,0,0,0,4,0,6,0,0,0,0,0,0,0,40,1
Oliveira,3,1,3,0,6,0,3,0,8,0,0,0,0,0,5,0,9,0,3,0,0,0,0,0,40,1
Gamarra,0,0,12,1,12,0,4,0,8,0,0,0,0,0,1,0,0,0,0,0,0,0,0,0,37,1
I. Galvão,0,0,1,0,4,0,3,0,8,0,4,0,0,0,7,0,10,0,0,0,0,0,0,0,37,0
Rapha Diniz,0,0,10,2,4,0,0,0,3,0,6,0,0,0,4,0,7,0,3,0,0,0,0,0,37,2
Vanderson,0,0,3,1,10,0,8,0,2,0,6,0,0,0,7,0,0,0,0,0,0,0,0,0,36,1
Morais,3,1,9,1,3,0,4,0,9,0,4,0,0,0,0,0,4,0,0,0,0,0,0,0,36,2
André,0,0,0,0,4,0,4,0,13,0,1,0,0,0,6,0,6,0,0,0,0,0,0,0,34,0
R. Mateus,3,1,4,0,7,0,0,0,5,0,4,0,0,0,5,0,6,0,0,0,0,0,0,0,34,1
Michel,0,0,0,0,1,0,3,0,11,0,4,0,0,0,5,0,7,0,0,0,0,0,0,0,31,0
Garcia,0,0,7,2,6,0,4,0,8,0,0,0,0,0,0,0,3,0,3,0,0,0,0,0,31,2
Marcelinho,3,1,1,0,16,0,10,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,30,1
Harley,0,0,3,0,4,0,4,0,5,0,4,0,0,0,4,0,3,0,3,0,0,0,0,0,30,0
Otaner,0,0,3,0,8,0,1,0,1,0,3,0,0,0,7,0,3,0,0,0,0,0,0,0,26,0
Gusta,0,0,0,0,9,0,0,0,0,0,3,0,0,0,5,0,9,0,0,0,0,0,0,0,26,0
M. Galvão,3,1,4,0,4,0,4,0,0,0,0,0,0,0,10,0,0,0,0,0,0,0,0,0,25,1
Filipinho,0,0,3,1,0,0,0,0,4,0,0,0,0,0,9,0,9,0,0,0,0,0,0,0,25,1
Murilo,0,0,0,0,0,0,5,0,5,0,0,0,0,0,9,0,3,0,3,0,0,0,0,0,25,0
Gutto,0,0,0,0,4,0,5,0,5,0,0,0,0,0,0,0,10,0,0,0,0,0,0,0,24,0
Nando,0,0,1,0,0,0,0,0,3,0,0,0,0,0,5,0,12,0,3,0,0,0,0,0,24,0
Diniz,0,0,7,1,0,0,3,0,6,0,6,0,0,0,1,0,0,0,0,0,0,0,0,0,23,1
Dyeggo,0,0,7,0,3,0,0,0,0,0,0,0,0,0,0,0,10,0,0,0,0,0,0,0,20,0
Carioca,0,0,1,0,0,0,6,0,4,0,3,0,0,0,4,0,0,0,0,0,0,0,0,0,18,0
Zé Paulo,0,0,3,0,6,0,3,0,3,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0,16,0
Juracy,0,0,0,0,6,0,7,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,13,0
Kiko,3,1,4,1,0,0,0,0,3,0,3,0,0,0,0,0,0,0,0,0,0,0,0,0,13,2
PH,0,0,1,0,3,0,0,0,0,0,7,0,0,0,0,0,0,0,0,0,0,0,0,0,11,0
Babinha,0,0,0,0,6,0,0,0,1,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0,8,0
Guga,0,0,0,0,4,0,0,0,0,0,0,0,0,0,0,0,3,0,0,0,0,0,0,0,7,0
Marcola,0,0,0,0,0,0,0,0,0,0,0,0,0,0,3,0,4,0,0,0,0,0,0,0,7,0
Miskinis,0,0,0,0,0,0,3,0,0,0,0,0,0,0,3,0,0,0,0,0,0,0,0,0,6,0
Dai,0,0,0,0,3,0,0,0,0,0,0,0,0,0,1,0,0,0,0,0,0,0,0,0,4,0
Gilberto,0,0,3,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,3,1
Jefinho,0,0,0,0,0,0,3,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,3,0
Beradeiro,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0
Ivamário,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0
Jonathan,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"""

def parse_raw_dataframe(df_raw: pd.DataFrame) -> pd.DataFrame:
    df = df_raw.copy()
    first_col = df.columns[0]
    df.rename(columns={first_col: "Jogador"}, inplace=True)
    df = df[df["Jogador"].notna()]
    df = df[~df["Jogador"].astype(str).str.startswith("Sábados")]
    df = df[~df["Jogador"].astype(str).str.startswith("Coluna")]
    df["Jogador"] = df["Jogador"].astype(str).str.strip()
    
    num_cols = [c for c in df.columns if c != "Jogador"]
    for c in num_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
    return df

@st.cache_data
def load_all_sheets(uploaded_file=None, file_path="Prêmios 2026.xlsx"):
    sheets = {}
    target = uploaded_file if uploaded_file is not None else (file_path if os.path.exists(file_path) else None)
    expected_sheets = ["Planilha Craque", "Planilha Brocador", "Planilha César", "Planilha Reis do Sábado"]
    
    if target is not None:
        try:
            excel = pd.ExcelFile(target)
            available = excel.sheet_names
            for sheet in expected_sheets:
                if sheet in available:
                    df_temp = pd.read_excel(excel, sheet_name=sheet)
                    if "Coluna1" in df_temp.columns or "Prêmio" in str(df_temp.iloc[0, 0]):
                        df_temp = pd.read_excel(excel, sheet_name=sheet, header=1)
                    sheets[sheet] = parse_raw_dataframe(df_temp)
        except Exception as e:
            st.sidebar.warning(f"Aviso ao ler Excel: {e}")

    base_df = parse_raw_dataframe(pd.read_csv(io.StringIO(RAW_DATA_CSV)))
    
    if "Planilha Craque" not in sheets or sheets["Planilha Craque"] is None:
        sheets["Planilha Craque"] = base_df.copy()
    if "Planilha Brocador" not in sheets or sheets["Planilha Brocador"] is None:
        df_brocador = base_df.copy()
        for m in MONTHS_LIST:
            df_brocador[m] = (df_brocador[m] * 0.45).round().astype(int)
        df_brocador["Pontuação Geral"] = df_brocador[MONTHS_LIST].sum(axis=1)
        sheets["Planilha Brocador"] = df_brocador.sort_values(by="Pontuação Geral", ascending=False)
    if "Planilha César" not in sheets or sheets["Planilha César"] is None:
        goleiros = ["Robério", "Garcia", "Remi", "Albino", "Babinha", "Diniz", "Gamarra", "Canário", "Jandir", "Adriano"]
        df_cesar = base_df[base_df["Jogador"].isin(goleiros)].copy()
        for m in MONTHS_LIST:
            df_cesar[m] = (df_cesar[m] * 0.75 + 1).round().astype(int)
        df_cesar["Pontuação Geral"] = df_cesar[MONTHS_LIST].sum(axis=1)
        sheets["Planilha César"] = df_cesar.sort_values(by="Pontuação Geral", ascending=False)
    if "Planilha Reis do Sábado" not in sheets or sheets["Planilha Reis do Sábado"] is None:
        sheets["Planilha Reis do Sábado"] = base_df.copy().sort_values(by=["RS Total", "Pontuação Geral"], ascending=[False, False])

    return sheets

# ======================================================================================
# SIDEBAR
# ======================================================================================
st.sidebar.title("⚽ Prêmio Mangabeira")
st.sidebar.caption("Campeonato de Futebol 2026")

uploaded_file = st.sidebar.file_uploader("Carregar Prêmios 2026.xlsx", type=["xlsx", "xls", "csv"])
sheets_dict = load_all_sheets(uploaded_file)

categoria_selecionada = st.sidebar.selectbox(
    "Aba da Planilha:",
    options=["Planilha Craque", "Planilha Brocador", "Planilha César", "Planilha Reis do Sábado"],
    index=0
)

meses_opcoes = ["Geral (Acumulado)"] + MONTHS_LIST
mes_selecionado = st.sidebar.selectbox("Filtro de Mês:", meses_opcoes, index=0)
top_n = st.sidebar.slider("Top Jogadores no Gráfico:", min_value=5, max_value=25, value=10)
busca_jogador = st.sidebar.text_input("🔍 Buscar Jogador:", placeholder="Digite o nome...")

# ======================================================================================
# DADOS SELECIONADOS
# ======================================================================================
df_atual = sheets_dict[categoria_selecionada].copy()
is_geral = (mes_selecionado == "Geral (Acumulado)")

if is_geral:
    col_pontos = "Pontuação Geral"
    col_rs = "RS Total" if "RS Total" in df_atual.columns else None
    titulo_periodo = "Classificação Geral (Temporada Completa)"
else:
    col_pontos = mes_selecionado
    mes_idx = MONTHS_LIST.index(mes_selecionado) + 1
    col_rs = f"RS {mes_idx}" if f"RS {mes_idx}" in df_atual.columns else None
    titulo_periodo = f"Desempenho no mês de {mes_selecionado}"

if categoria_selecionada == "Planilha Reis do Sábado" and is_geral:
    df_sorted = df_atual.sort_values(by=["RS Total", col_pontos], ascending=[False, False]).reset_index(drop=True)
else:
    df_sorted = df_atual.sort_values(by=col_pontos, ascending=False).reset_index(drop=True)

if busca_jogador:
    df_filtrado = df_sorted[df_sorted["Jogador"].str.contains(busca_jogador, case=False, na=False)].reset_index(drop=True)
else:
    df_filtrado = df_sorted

icones_cat = {
    "Planilha Craque": ("🏆", "Craque do Ano (Melhor Jogador)", "pts"),
    "Planilha Brocador": ("⚽", "Prêmio Brocador (Artilharia)", "gols"),
    "Planilha César": ("🧤", "Prêmio César (Melhor Goleiro)", "pts"),
    "Planilha Reis do Sábado": ("👑", "Reis do Sábado (Coroas Individuais)", "títulos")
}
icone_cat, nome_cat, unit = icones_cat.get(categoria_selecionada, ("🏆", categoria_selecionada, "pts"))

# ======================================================================================
# CABEÇALHO & DESTAQUES DOS LÍDERES
# ======================================================================================
st.title(f"{icone_cat} {nome_cat}")
st.caption(f"Visualizando: **{titulo_periodo}** · Total de atletas: **{len(df_atual)}**")
st.markdown("---")

p1 = df_sorted.iloc[0] if len(df_sorted) > 0 else None
p2 = df_sorted.iloc[1] if len(df_sorted) > 1 else None
p3 = df_sorted.iloc[2] if len(df_sorted) > 2 else None
df_rs_leader = df_atual.sort_values(by="RS Total", ascending=False).iloc[0] if "RS Total" in df_atual.columns else p1
media_pts = df_sorted[col_pontos].mean()

c1, c2, c3, c4, c5 = st.columns(5)
with c1:
    if p1 is not None:
        st.markdown(f"""
        <div class="metric-card" style="border-top: 3px solid #fbbf24;">
            <div class="metric-rank rank-gold">🥇 1º Lugar (Líder)</div>
            <div class="metric-player">{p1['Jogador']}</div>
            <div class="metric-value">{p1[col_pontos]} <span style="font-size:0.9rem;color:#9ca3af;">{unit}</span></div>
            <div class="metric-sub">RS no Período: {p1[col_rs] if col_rs else '-'}</div>
        </div>
        """, unsafe_allow_html=True)
with c2:
    if p2 is not None:
        st.markdown(f"""
        <div class="metric-card" style="border-top: 3px solid #94a3b8;">
            <div class="metric-rank rank-silver">🥈 2º Lugar</div>
            <div class="metric-player">{p2['Jogador']}</div>
            <div class="metric-value" style="color:#e2e8f0;">{p2[col_pontos]} <span style="font-size:0.9rem;color:#9ca3af;">{unit}</span></div>
            <div class="metric-sub">Diferença: -{p1[col_pontos] - p2[col_pontos]} {unit}</div>
        </div>
        """, unsafe_allow_html=True)
with c3:
    if p3 is not None:
        st.markdown(f"""
        <div class="metric-card" style="border-top: 3px solid #f97316;">
            <div class="metric-rank rank-bronze">🥉 3º Lugar</div>
            <div class="metric-player">{p3['Jogador']}</div>
            <div class="metric-value" style="color:#fdba74;">{p3[col_pontos]} <span style="font-size:0.9rem;color:#9ca3af;">{unit}</span></div>
            <div class="metric-sub">Diferença: -{p1[col_pontos] - p3[col_pontos]} {unit}</div>
        </div>
        """, unsafe_allow_html=True)
with c4:
    if df_rs_leader is not None:
        st.markdown(f"""
        <div class="metric-card" style="border-top: 3px solid #ec4899;">
            <div class="metric-rank rank-crown">👑 Maior Rei do Sábado</div>
            <div class="metric-player">{df_rs_leader['Jogador']}</div>
            <div class="metric-value" style="color:#f472b6;">{df_rs_leader['RS Total']} <span style="font-size:0.9rem;color:#9ca3af;">títulos</span></div>
            <div class="metric-sub">Pontos Totais: {df_rs_leader['Pontuação Geral']} pts</div>
        </div>
        """, unsafe_allow_html=True)
with c5:
    st.markdown(f"""
    <div class="metric-card" style="border-top: 3px solid #38bdf8;">
        <div class="metric-rank rank-stats">📊 Média Geral</div>
        <div class="metric-player">Atletas Ativos</div>
        <div class="metric-value" style="color:#38bdf8;">{media_pts:.1f} <span style="font-size:0.9rem;color:#9ca3af;">{unit}</span></div>
        <div class="metric-sub">Total de {len(df_sorted)} inscritos</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")

# ======================================================================================
# GRÁFICOS (MATPLOTLIB & SEABORN)
# ======================================================================================
st.subheader(f"📈 Gráficos de Desempenho (Top {top_n})")
col_chart1, col_chart2 = st.columns([3, 2])

with col_chart1:
    st.markdown(f"##### Ranking dos Melhores Atletas ({titulo_periodo})")
    df_top = df_sorted.head(top_n).sort_values(by=col_pontos, ascending=True)
    
    fig, ax = plt.subplots(figsize=(10, max(5, top_n * 0.42)))
    fig.patch.set_facecolor('#0e1117')
    ax.set_facecolor('#111827')
    
    colors = sns.color_palette("mako", n_colors=len(df_top))
    bars = ax.barh(df_top["Jogador"], df_top[col_pontos], color=colors, height=0.68, edgecolor='#1f2937')
    
    for bar in bars:
        width = bar.get_width()
        ax.text(
            width + (max(df_top[col_pontos]) * 0.015),
            bar.get_y() + bar.get_height() / 2,
            f"{int(width)} {unit}",
            va='center', ha='left', color='#ffffff', fontsize=9.5, fontweight='bold'
        )
        
    ax.set_xlabel(f"Pontos ({unit})", fontsize=11, color='#9ca3af', labelpad=8)
    ax.set_xlim(0, max(df_top[col_pontos]) * 1.15 if max(df_top[col_pontos]) > 0 else 10)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.tick_params(colors='#e5e7eb', labelsize=10)
    ax.grid(axis='x', color='#1f2937', linestyle='--', alpha=0.7)
    plt.tight_layout()
    st.pyplot(fig)
    plt.close()

with col_chart2:
    st.markdown("##### 🚀 Evolução Mensal dos Top 5 Atletas")
    top5_players = df_sorted.head(5)["Jogador"].tolist()
    df_top5_months = df_atual[df_atual["Jogador"].isin(top5_players)]
    
    fig2, ax2 = plt.subplots(figsize=(8, 6.2))
    fig2.patch.set_facecolor('#0e1117')
    ax2.set_facecolor('#111827')
    
    palette = ["#10b981", "#fbbf24", "#38bdf8", "#ec4899", "#a855f7"]
    meses_com_jogos = [m for m in MONTHS_LIST if df_atual[m].sum() > 0] or MONTHS_LIST[:6]
        
    for idx, (_, row) in enumerate(df_top5_months.iterrows()):
        pts_history = [row[m] for m in meses_com_jogos]
        ax2.plot(meses_com_jogos, pts_history, marker='o', linewidth=2.4, label=row["Jogador"], color=palette[idx % len(palette)])
        
    ax2.set_ylabel(f"Pontos ({unit})", fontsize=10, color='#9ca3af')
    ax2.tick_params(colors='#e5e7eb', labelsize=9)
    plt.xticks(rotation=45, ha='right')
    ax2.spines['top'].set_visible(False)
    ax2.spines['right'].set_visible(False)
    ax2.grid(True, color='#1f2937', linestyle='--', alpha=0.7)
    ax2.legend(facecolor='#111827', edgecolor='#374151', labelcolor='#f3f4f6', fontsize=9)
    plt.tight_layout()
    st.pyplot(fig2)
    plt.close()

st.markdown("---")

# ======================================================================================
# TABELA COM FILTROS E EXPORTAÇÃO CSV
# ======================================================================================
st.subheader("📋 Tabela de Classificação Detalhada")

df_view = df_filtrado.copy()
df_view.insert(0, "Posição", [f"🥇 1º" if i == 1 else f"🥈 2º" if i == 2 else f"🥉 3º" if i == 3 else f"{i}º" for i in range(1, len(df_view) + 1)])

if is_geral:
    cols_to_show = ["Posição", "Jogador", "Pontuação Geral", "RS Total"] + [m for m in MONTHS_LIST if m in df_view.columns]
else:
    cols_to_show = ["Posição", "Jogador", col_pontos] + ([col_rs] if col_rs and col_rs in df_view.columns else []) + ["Pontuação Geral", "RS Total"]

df_table = df_view[[c for c in cols_to_show if c in df_view.columns]]

col_tab1, col_tab2 = st.columns([4, 1])
with col_tab1:
    st.dataframe(df_table, use_container_width=True, hide_index=True, height=450)
with col_tab2:
    st.markdown("##### 📥 Exportar")
    csv_data = df_table.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="Baixar Tabela (CSV)",
        data=csv_data,
        file_name=f"{categoria_selecionada.replace(' ', '_')}_{mes_selecionado}.csv",
        mime="text/csv",
        use_container_width=True
    )
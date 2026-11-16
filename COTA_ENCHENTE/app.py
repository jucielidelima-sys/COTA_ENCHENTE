from io import BytesIO
from pathlib import Path
import re

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

DEFAULT_BASE_PATH = Path(__file__).parent / "base" / "COTA_ENCHENTE.xlsx"

st.set_page_config(
    page_title="Dashboard Risco de Enchente",
    page_icon="🌊",
    layout="wide",
    initial_sidebar_state="expanded",
)

CUSTOM_CSS = """
<style>
    .stApp { background: linear-gradient(135deg, #061522 0%, #0a1f33 52%, #071522 100%); }
    div[data-testid="stMetric"] {
        background: linear-gradient(145deg, rgba(13,34,56,.95), rgba(10,27,45,.95));
        border: 1px solid rgba(40,241,198,.20);
        border-radius: 16px;
        padding: 18px 18px 14px 18px;
        box-shadow: 0 0 22px rgba(40,241,198,.08);
    }
    div[data-testid="stMetric"] label { color: #90a8bd !important; }
    div[data-testid="stMetricValue"] { color: #eaf4ff !important; }
    .block-container { padding-top: 1.4rem; }
    .painel-title {
        padding: 18px 20px;
        border-radius: 18px;
        background: linear-gradient(90deg, rgba(40,241,198,.16), rgba(65,105,225,.10));
        border: 1px solid rgba(40,241,198,.24);
        margin-bottom: 16px;
    }
    .painel-title h1 { margin: 0; font-size: 34px; color: #eaf4ff; }
    .painel-title p { margin: 6px 0 0 0; color: #9fb5c8; }
    .risk-card {
        background: rgba(13,34,56,.72);
        border: 1px solid rgba(255,255,255,.08);
        border-radius: 16px;
        padding: 14px 16px;
        margin-bottom: 8px;
    }
    .priority-box {
        background: rgba(13,34,56,.80);
        border: 1px solid rgba(40,241,198,.20);
        border-radius: 14px;
        padding: 12px 14px;
        margin-bottom: 8px;
        min-height: 82px;
    }
    .priority-box b { color: #eaf4ff; font-size: 18px; }
    .priority-box span { color: #9fb5c8; font-size: 13px; }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


def normalize_text(value: object) -> str:
    if pd.isna(value):
        return "NÃO INFORMADO"
    text = str(value).strip().upper()
    text = re.sub(r"\s+", " ", text)
    replacements = {"NAO": "NÃO", "APE": "APÉ", "A PÉ": "APÉ", "APÉ": "APÉ"}
    return replacements.get(text, text)


def parse_cota(value: object) -> float | None:
    if pd.isna(value):
        return None
    text = str(value).strip().upper().replace(",", ".")
    if text in {"N/P", "NP", "N.P", "NÃO PASSA", "NAO PASSA", "NÃO INFORMADO", ""}:
        return None
    match = re.search(r"(\d+(?:\.\d+)?)", text)
    return float(match.group(1)) if match else None


def localizar_base_excel() -> Path | None:
    """Localiza automaticamente a planilha dentro do projeto, mesmo no Streamlit Cloud."""
    base_dir = Path(__file__).resolve().parent
    candidatos = [
        DEFAULT_BASE_PATH,
        base_dir / "base" / "COTA ENCHENTE.xlsx",
        base_dir / "COTA_ENCHENTE.xlsx",
        base_dir / "COTA ENCHENTE.xlsx",
        Path.cwd() / "base" / "COTA_ENCHENTE.xlsx",
        Path.cwd() / "base" / "COTA ENCHENTE.xlsx",
    ]
    for caminho in candidatos:
        if caminho.exists():
            return caminho

    # Busca em pastas próximas para evitar erro quando o GitHub/Streamlit cria subpastas.
    raizes = [base_dir, Path.cwd(), base_dir.parent, base_dir.parent.parent]
    for raiz in raizes:
        if raiz.exists():
            arquivos = list(raiz.glob("**/COTA*ENCHENTE*.xlsx")) + list(raiz.glob("**/*ENCHENTE*.xlsx"))
            arquivos = [a for a in arquivos if a.is_file() and not a.name.startswith("~$")]
            if arquivos:
                return arquivos[0]
    return None


@st.cache_data(show_spinner=False)
def load_data(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path)
    df.columns = [str(c).strip().upper() for c in df.columns]

    rename_map = {
        "NOME DO FUNCIONÁRIO": "COLABORADOR",
        "NOME DO FUNCIONARIO": "COLABORADOR",
        "COTA CHEIA": "COTA CHEIA",
        "ACESSO A EMPRESA": "ACESSO ATÉ EMPRESA",
        "ACESSO ATÉ EMPRESA": "ACESSO ATÉ EMPRESA",
        "MEIO DE TRANSPORTE": "MEIO TRANSPORTE",
        "MEIO TRANSPORTE": "MEIO TRANSPORTE",
        "ÁREA": "SETOR",
        "AREA": "SETOR",
        "DEPARTAMENTO": "SETOR",
    }
    df = df.rename(columns={c: rename_map.get(c, c) for c in df.columns})

    required = ["COLABORADOR", "SETOR", "COTA CHEIA", "ACESSO ATÉ EMPRESA", "MEIO TRANSPORTE"]
    for col in required:
        if col not in df.columns:
            df[col] = "NÃO INFORMADO"

    df = df[required].copy()
    df["COLABORADOR"] = df["COLABORADOR"].fillna("NÃO INFORMADO").astype(str).str.strip()
    df["SETOR"] = df["SETOR"].map(normalize_text)
    df["ACESSO ATÉ EMPRESA"] = df["ACESSO ATÉ EMPRESA"].map(normalize_text)
    df["MEIO TRANSPORTE"] = df["MEIO TRANSPORTE"].map(normalize_text)
    df["COTA ORIGINAL"] = df["COTA CHEIA"].astype(str).str.strip().str.upper()
    df["COTA NUM"] = df["COTA CHEIA"].map(parse_cota)
    df["STATUS COTA"] = df["COTA NUM"].apply(lambda x: "SEM PASSAGEM / N/P" if pd.isna(x) else f"{x:g}M")
    return df


def to_excel_bytes(data: pd.DataFrame) -> bytes:
    """Gera Excel sem depender do pacote xlsxwriter."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        data.to_excel(writer, index=False, sheet_name="Dados Filtrados")
        ws = writer.book["Dados Filtrados"]
        for idx, col in enumerate(data.columns, start=1):
            serie = data[col].astype(str)
            width = min(max(len(str(col)) + 2, int(serie.str.len().quantile(0.9)) + 2), 42)
            ws.column_dimensions[ws.cell(row=1, column=idx).column_letter].width = width
    return output.getvalue()


def calcula_impacto(cota_num: object, acesso: object, nivel: float) -> str:
    """
    Regra para o gráfico Pessoas Impactadas × Nível do Rio:
    - Cota numérica: impacta quando o nível do rio é maior ou igual à cota do colaborador.
    - N/P ou sem cota: só entra como impactado quando o acesso até a empresa está como NÃO.
    - O campo Acesso até empresa continua aparecendo como indicador separado, mas não zera a contagem de uma cota já atingida.
    """
    if not pd.isna(cota_num):
        return "SIM" if float(cota_num) <= float(nivel) else "NÃO"
    return "SIM" if normalize_text(acesso) == "NÃO" else "NÃO"




def calcula_prioridade(cota_num: object, acesso: object, nivel: float) -> str:
    """Classifica a prioridade operacional por colaborador.

    P1: cota já atingida pelo nível atual, ou N/P com acesso NÃO.
    P2: cota até 1m acima do nível atual, ou acesso NÃO com cota ainda não atingida.
    P3: cota até 2m acima do nível atual.
    P4: rotina/sem ação imediata.
    """
    acesso_nao = normalize_text(acesso) == "NÃO"
    if pd.isna(cota_num):
        return "P1 - AÇÃO IMEDIATA" if acesso_nao else "P4 - ROTINA"

    margem = float(cota_num) - float(nivel)
    if margem <= 0:
        return "P1 - AÇÃO IMEDIATA"
    if margem <= 1 or acesso_nao:
        return "P2 - ALTA ATENÇÃO"
    if margem <= 2:
        return "P3 - MONITORAR"
    return "P4 - ROTINA"


def matriz_prioridade(data: pd.DataFrame) -> pd.DataFrame:
    matriz = (
        data.groupby(["SETOR", "PRIORIDADE OPERACIONAL"], as_index=False)
        .size()
        .rename(columns={"size": "QTD"})
    )
    ordem = {
        "P1 - AÇÃO IMEDIATA": 1,
        "P2 - ALTA ATENÇÃO": 2,
        "P3 - MONITORAR": 3,
        "P4 - ROTINA": 4,
    }
    matriz["ORDEM"] = matriz["PRIORIDADE OPERACIONAL"].map(ordem).fillna(9)
    return matriz.sort_values(["ORDEM", "SETOR"])


def ranking_operacional(data: pd.DataFrame) -> pd.DataFrame:
    base = data.copy()
    base["P1"] = (base["PRIORIDADE OPERACIONAL"] == "P1 - AÇÃO IMEDIATA").astype(int)
    base["P2"] = (base["PRIORIDADE OPERACIONAL"] == "P2 - ALTA ATENÇÃO").astype(int)
    base["P3"] = (base["PRIORIDADE OPERACIONAL"] == "P3 - MONITORAR").astype(int)
    base["P4"] = (base["PRIORIDADE OPERACIONAL"] == "P4 - ROTINA").astype(int)
    base["SEM ACESSO"] = (base["ACESSO ATÉ EMPRESA"] == "NÃO").astype(int)
    rank = (
        base.groupby("SETOR", as_index=False)
        .agg(
            COLABORADORES=("COLABORADOR", "count"),
            P1=("P1", "sum"),
            P2=("P2", "sum"),
            P3=("P3", "sum"),
            P4=("P4", "sum"),
            SEM_ACESSO=("SEM ACESSO", "sum"),
        )
    )
    rank["SCORE"] = rank["P1"] * 100 + rank["P2"] * 60 + rank["P3"] * 30 + rank["SEM_ACESSO"] * 10
    rank["% IMPACTO"] = ((rank["P1"] / rank["COLABORADORES"]) * 100).round(1)
    return rank.sort_values(["SCORE", "P1", "P2"], ascending=False)


def acao_recomendada(prioridade: str) -> str:
    mapa = {
        "P1 - AÇÃO IMEDIATA": "Acionar liderança, confirmar presença, rota alternativa ou dispensa preventiva.",
        "P2 - ALTA ATENÇÃO": "Manter contato ativo, validar transporte e preparar plano de contingência.",
        "P3 - MONITORAR": "Acompanhar evolução do nível do rio e revisar a cada atualização.",
        "P4 - ROTINA": "Sem ação imediata, manter no monitoramento geral.",
    }
    return mapa.get(prioridade, "Monitorar.")

def curva_impacto(data: pd.DataFrame) -> pd.DataFrame:
    cotas = sorted(data["COTA NUM"].dropna().unique())
    linhas = []
    for nivel in cotas:
        qtd = data.apply(lambda r: calcula_impacto(r["COTA NUM"], r["ACESSO ATÉ EMPRESA"], nivel), axis=1).eq("SIM").sum()
        linhas.append({"NÍVEL DO RIO (m)": float(nivel), "PESSOAS IMPACTADAS": int(qtd)})
    return pd.DataFrame(linhas)


base_encontrada = localizar_base_excel()
if base_encontrada is None:
    st.error("Base Excel não encontrada. Confirme se o arquivo está dentro da pasta base/ com o nome COTA_ENCHENTE.xlsx.")
    st.stop()

df = load_data(base_encontrada)

st.markdown(
    """
    <div class="painel-title">
        <h1>🌊 Dashboard de Risco de Enchente</h1>
        <p>Monitoramento de colaboradores impactados por cota, setor, acesso à empresa e meio de transporte.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("⚙️ Filtros")
    cotas_validas = sorted(df["COTA NUM"].dropna().unique())
    min_cota = float(min(cotas_validas)) if cotas_validas else 0.0
    max_cota = float(max(cotas_validas)) if cotas_validas else 15.0
    nivel_atual = st.slider("Nível atual da enchente (m)", min_value=min_cota, max_value=max_cota, value=min_cota, step=0.5)

    setores = sorted(df["SETOR"].dropna().unique())
    setor_sel = st.multiselect("Setor", setores, default=setores)

    transportes = sorted(df["MEIO TRANSPORTE"].dropna().unique())
    transp_sel = st.multiselect("Meio de transporte", transportes, default=transportes)

    acessos = sorted(df["ACESSO ATÉ EMPRESA"].dropna().unique())
    acesso_sel = st.multiselect("Acesso até empresa", acessos, default=acessos)

    busca = st.text_input("Pesquisar colaborador")

filtered = df[
    df["SETOR"].isin(setor_sel)
    & df["MEIO TRANSPORTE"].isin(transp_sel)
    & df["ACESSO ATÉ EMPRESA"].isin(acesso_sel)
].copy()
if busca:
    filtered = filtered[filtered["COLABORADOR"].str.contains(busca, case=False, na=False)]

filtered["IMPACTADO NA COTA ATUAL"] = filtered.apply(
    lambda r: calcula_impacto(r["COTA NUM"], r["ACESSO ATÉ EMPRESA"], nivel_atual), axis=1
)
filtered["RISCO"] = filtered.apply(
    lambda r: "CRÍTICO" if r["IMPACTADO NA COTA ATUAL"] == "SIM" else "MONITORADO",
    axis=1,
)
filtered["PRIORIDADE OPERACIONAL"] = filtered.apply(
    lambda r: calcula_prioridade(r["COTA NUM"], r["ACESSO ATÉ EMPRESA"], nivel_atual), axis=1
)
filtered["AÇÃO RECOMENDADA"] = filtered["PRIORIDADE OPERACIONAL"].map(acao_recomendada)

colab_total = len(filtered)
impactados = int((filtered["IMPACTADO NA COTA ATUAL"] == "SIM").sum())
sem_acesso = int((filtered["ACESSO ATÉ EMPRESA"] == "NÃO").sum())
setores_criticos = int(filtered.loc[filtered["RISCO"] == "CRÍTICO", "SETOR"].nunique())
perc_impacto = (impactados / colab_total * 100) if colab_total else 0
p1_total = int((filtered["PRIORIDADE OPERACIONAL"] == "P1 - AÇÃO IMEDIATA").sum())
p2_total = int((filtered["PRIORIDADE OPERACIONAL"] == "P2 - ALTA ATENÇÃO").sum())

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Colaboradores", f"{colab_total}")
k2.metric("Impactados na cota", f"{impactados}", f"{perc_impacto:.1f}%")
k3.metric("Sem acesso", f"{sem_acesso}")
k4.metric("Setores críticos", f"{setores_criticos}")
k5.metric("Nível simulado", f"{nivel_atual:g} m")

left, right = st.columns([1.1, 1])
with left:
    gauge = go.Figure(go.Indicator(
        mode="gauge+number+delta",
        value=perc_impacto,
        number={"suffix": "%"},
        title={"text": "Impacto na Cota Atual"},
        delta={"reference": 30, "suffix": "% ref."},
        gauge={
            "axis": {"range": [0, 100]},
            "bar": {"color": "#28f1c6"},
            "steps": [
                {"range": [0, 20], "color": "rgba(30,160,120,.28)"},
                {"range": [20, 50], "color": "rgba(230,190,60,.30)"},
                {"range": [50, 100], "color": "rgba(230,70,80,.34)"},
            ],
        },
    ))
    gauge.update_layout(height=330, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font_color="#eaf4ff")
    st.plotly_chart(gauge, use_container_width=True)

with right:
    st.markdown("#### 🚦 Leitura operacional")
    st.markdown(
        f"""
        <div class="risk-card">Nível simulado: <b>{nivel_atual:g} m</b></div>
        <div class="risk-card">Colaboradores atingidos pela cota atual: <b>{impactados}</b></div>
        <div class="risk-card">Colaboradores com acesso informado como NÃO: <b>{sem_acesso}</b></div>
        <div class="risk-card">Setores com algum risco crítico: <b>{setores_criticos}</b></div>
        """,
        unsafe_allow_html=True,
    )

c1, c2 = st.columns(2)
with c1:
    setor_chart = filtered.groupby(["SETOR", "RISCO"], as_index=False).size().rename(columns={"size": "QTD"})
    fig = px.bar(setor_chart, x="SETOR", y="QTD", color="RISCO", title="Risco por Setor", text="QTD")
    fig.update_layout(height=430, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font_color="#eaf4ff", xaxis_tickangle=-35)
    st.plotly_chart(fig, use_container_width=True)

with c2:
    impacto_nivel = curva_impacto(filtered)
    fig = px.line(
        impacto_nivel,
        x="NÍVEL DO RIO (m)",
        y="PESSOAS IMPACTADAS",
        markers=True,
        text="PESSOAS IMPACTADAS",
        title="Quantidade de Pessoas Impactadas × Nível do Rio",
    )
    fig.update_traces(textposition="top center")
    fig.update_layout(height=430, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font_color="#eaf4ff")
    st.plotly_chart(fig, use_container_width=True)

st.markdown("### 🧭 Matriz de Prioridade Operacional")

pcol1, pcol2, pcol3, pcol4 = st.columns(4)
pcol1.markdown(f'<div class="priority-box"><b>{p1_total}</b><br><span>P1 - Ação imediata</span></div>', unsafe_allow_html=True)
pcol2.markdown(f'<div class="priority-box"><b>{p2_total}</b><br><span>P2 - Alta atenção</span></div>', unsafe_allow_html=True)
pcol3.markdown(f'<div class="priority-box"><b>{int((filtered["PRIORIDADE OPERACIONAL"] == "P3 - MONITORAR").sum())}</b><br><span>P3 - Monitorar</span></div>', unsafe_allow_html=True)
pcol4.markdown(f'<div class="priority-box"><b>{int((filtered["PRIORIDADE OPERACIONAL"] == "P4 - ROTINA").sum())}</b><br><span>P4 - Rotina</span></div>', unsafe_allow_html=True)

pmat = matriz_prioridade(filtered)
rank = ranking_operacional(filtered)

mp1, mp2 = st.columns([1.1, 1])
with mp1:
    if not pmat.empty:
        pivot = pmat.pivot_table(index="SETOR", columns="PRIORIDADE OPERACIONAL", values="QTD", aggfunc="sum", fill_value=0)
        ordem_cols = ["P1 - AÇÃO IMEDIATA", "P2 - ALTA ATENÇÃO", "P3 - MONITORAR", "P4 - ROTINA"]
        pivot = pivot.reindex(columns=[c for c in ordem_cols if c in pivot.columns], fill_value=0)
        fig = px.imshow(
            pivot,
            text_auto=True,
            aspect="auto",
            title="Matriz Setor × Prioridade",
            labels=dict(x="Prioridade", y="Setor", color="Qtd"),
        )
        fig.update_layout(height=460, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font_color="#eaf4ff")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Sem dados para montar a matriz de prioridade.")

with mp2:
    st.markdown("#### Ranking de acionamento por setor")
    cols_rank = ["SETOR", "COLABORADORES", "P1", "P2", "P3", "SEM_ACESSO", "% IMPACTO", "SCORE"]
    st.dataframe(rank[cols_rank], use_container_width=True, hide_index=True)

st.markdown("#### Regras da matriz")
st.caption("P1: impacto imediato na cota atual ou N/P com acesso NÃO. P2: cota até 1m acima do nível atual e acesso NÃO. P3: cota até 2m acima do nível atual e acesso NÃO. P4: rotina/sem ação imediata.")

st.markdown("### 🌊 Distribuição por cota informada")
cota_chart = filtered.groupby("STATUS COTA", as_index=False).size().rename(columns={"size": "QTD"})
cota_chart["ORDEM"] = cota_chart["STATUS COTA"].map(lambda x: 999 if "N/P" in str(x) else float(str(x).replace("M", "")))
cota_chart = cota_chart.sort_values("ORDEM")
fig = px.bar(cota_chart, x="STATUS COTA", y="QTD", title="Quantidade de colaboradores por cota cadastrada", text="QTD")
fig.update_layout(height=360, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font_color="#eaf4ff")
st.plotly_chart(fig, use_container_width=True)

c3, c4 = st.columns(2)
with c3:
    transp = filtered.groupby("MEIO TRANSPORTE", as_index=False).size().rename(columns={"size": "QTD"})
    fig = px.pie(transp, names="MEIO TRANSPORTE", values="QTD", title="Meio de Transporte")
    fig.update_layout(height=390, paper_bgcolor="rgba(0,0,0,0)", font_color="#eaf4ff")
    st.plotly_chart(fig, use_container_width=True)

with c4:
    acesso = filtered.groupby("ACESSO ATÉ EMPRESA", as_index=False).size().rename(columns={"size": "QTD"})
    fig = px.pie(acesso, names="ACESSO ATÉ EMPRESA", values="QTD", title="Acesso até Empresa")
    fig.update_layout(height=390, paper_bgcolor="rgba(0,0,0,0)", font_color="#eaf4ff")
    st.plotly_chart(fig, use_container_width=True)

st.markdown("### 📋 Base filtrada")
show_cols = ["COLABORADOR", "SETOR", "COTA ORIGINAL", "ACESSO ATÉ EMPRESA", "MEIO TRANSPORTE", "IMPACTADO NA COTA ATUAL", "RISCO", "PRIORIDADE OPERACIONAL", "AÇÃO RECOMENDADA"]
st.dataframe(filtered[show_cols], use_container_width=True, hide_index=True)

st.download_button(
    "📥 Baixar base filtrada em Excel",
    data=to_excel_bytes(filtered[show_cols]),
    file_name="base_filtrada_risco_enchente.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)



with st.expander("ℹ️ Como o dashboard calcula o risco"):
    st.write(
        "O colaborador é classificado como impactado quando a cota cheia é menor ou igual ao nível do rio simulado. "
        "Quando a cota está como N/P e o acesso está como NÃO, o dashboard considera como sem passagem e impactado. "
        "A curva Pessoas Impactadas × Nível do Rio mostra o total acumulado, não apenas a quantidade exata daquela cota."
    )

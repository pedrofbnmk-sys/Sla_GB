import re
import unicodedata
from io import BytesIO
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="Indicador de SLA - Gestão de Bases",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================================================
# CONFIG
# =========================================================
COLUNAS_OBRIGATORIAS = [
    "Número de pedido JMS",
    "Base de entrega",
    "Responsável pela entrega",
    "Distrito destinatário",
    "Marca de assinatura",
    "Data de criação",
    "Horário de saída para entrega",
    "Horário da entrega",
    "Motivos dos pacotes problemáticos",
]

COLUNAS_OPCIONAIS = [
    "Cidade destinatária",
]

COLUNAS_LEITURA = COLUNAS_OBRIGATORIAS + COLUNAS_OPCIONAIS


# =========================================================
# UTILITÁRIOS
# =========================================================
def normalizar_texto_bruto(valor):
    if pd.isna(valor):
        return ""

    valor = str(valor).strip()
    if not valor:
        return ""

    valor = unicodedata.normalize("NFKD", valor)
    valor = "".join(ch for ch in valor if not unicodedata.combining(ch))
    valor = re.sub(r"\s+", " ", valor)
    return valor.upper().strip()


def titulo_suave(valor):
    if not valor:
        return "Não informado"

    valor = str(valor).strip().lower()
    valor = re.sub(r"\s+", " ", valor)
    return " ".join(p.capitalize() for p in valor.split(" "))


def padronizar_generico(valor):
    bruto = normalizar_texto_bruto(valor)
    if not bruto:
        return "Não informado"
    return titulo_suave(bruto)


def padronizar_bairro(valor):
    return padronizar_generico(valor)


def padronizar_cidade(valor):
    return padronizar_generico(valor)


def padronizar_entregador(valor):
    bruto = normalizar_texto_bruto(valor)
    if not bruto:
        return "Não informado"

    bruto = re.sub(r"\s*-\s*", " - ", bruto)
    bruto = re.sub(r"\s+", " ", bruto).strip()

    partes = bruto.split(" - ", 1)
    if len(partes) == 2:
        prefixo, nome = partes
        return f"{prefixo} - {titulo_suave(nome)}"

    return titulo_suave(bruto)


def padronizar_base(valor):
    bruto = normalizar_texto_bruto(valor)
    if not bruto:
        return "Não informado"
    bruto = re.sub(r"\s+", " ", bruto)
    return bruto


def classificar_status_sla(marca_assinatura):
    texto = normalizar_texto_bruto(marca_assinatura)

    if not texto:
        return "Outros"

    if "ASSINATURA DE DEVOLUCAO" in texto:
        return "Devolução"

    if "NAO ENTREGUE" in texto:
        return "Não entregue"

    if "RECEBIMENTO" in texto:
        return "Entregue"

    return "Outros"


def tratar_motivo(valor):
    if pd.isna(valor) or str(valor).strip() == "":
        return "Sem motivo informado"
    return str(valor).strip()


def validar_colunas(df):
    return [col for col in COLUNAS_OBRIGATORIAS if col not in df.columns]


def formatar_percentual(valor):
    return f"{valor:.2f}%"


def formatar_inteiro(valor):
    return f"{int(valor):,}".replace(",", ".")


def classificar_faixa_pendente(x):
    if pd.isna(x):
        return np.nan
    if x == 1:
        return "D1"
    if x == 2:
        return "D2"
    if x >= 3:
        return "D3+"
    return np.nan


def dataframe_to_excel_bytes(dfs_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for nome_aba, df in dfs_dict.items():
            df.to_excel(writer, sheet_name=str(nome_aba)[:31], index=False)
    output.seek(0)
    return output.getvalue()


# =========================================================
# META SLA
# =========================================================
def obter_meta_por_data(data_ref):
    # Monday=0 ... Sunday=6
    return 70.0 if pd.Timestamp(data_ref).weekday() == 6 else 96.0


def calcular_meta_periodo(data_inicio, data_fim):
    datas = pd.date_range(start=data_inicio, end=data_fim, freq="D")
    metas = [obter_meta_por_data(d) for d in datas]
    return float(np.mean(metas)) if metas else 96.0


def classificar_farol_sla(sla_realizado, meta_periodo):
    gap = sla_realizado - meta_periodo
    if gap >= 0:
        return "🟢 Acima da meta"
    elif gap >= -2:
        return "🟡 Próximo da meta"
    return "🔴 Abaixo da meta"


# =========================================================
# GRÁFICOS
# =========================================================
def criar_grafico_rosca_sla(sla_realizado, meta_periodo):
    valor = max(min(float(sla_realizado), 100.0), 0.0)
    restante = 100.0 - valor

    fig = go.Figure(
        data=[
            go.Pie(
                labels=["SLA realizado", "Restante"],
                values=[valor, restante],
                hole=0.72,
                sort=False,
                textinfo="none"
            )
        ]
    )

    fig.update_layout(
        margin=dict(t=20, b=20, l=20, r=20),
        showlegend=True,
        annotations=[
            dict(
                text=f"{valor:.2f}%",
                x=0.5,
                y=0.54,
                font_size=24,
                showarrow=False
            ),
            dict(
                text=f"Meta: {meta_periodo:.2f}%",
                x=0.5,
                y=0.40,
                font_size=12,
                showarrow=False
            )
        ]
    )
    return fig


def criar_grafico_linha_acumulada(df_baixas_hora):
    if df_baixas_hora.empty:
        return go.Figure()

    df_plot = df_baixas_hora.copy()
    df_plot["qtd_acumulada"] = df_plot["qtd_baixas"].cumsum()

    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=df_plot["hora"],
            y=df_plot["qtd_acumulada"],
            mode="lines+markers",
            name="Baixas acumuladas"
        )
    )
    fig.update_layout(
        xaxis_title="Hora",
        yaxis_title="Baixas acumuladas",
        margin=dict(t=20, b=20, l=20, r=20),
        legend_title=""
    )
    return fig


def criar_heatmap_entregador_hora(matriz):
    if matriz.empty:
        return go.Figure()

    df_heat = matriz.set_index("entregador_padronizado")
    z = df_heat.values
    x = list(df_heat.columns)
    y = list(df_heat.index)

    fig = go.Figure(
        data=go.Heatmap(
            z=z,
            x=x,
            y=y
        )
    )

    fig.update_layout(
        xaxis_title="Hora",
        yaxis_title="Entregador",
        margin=dict(t=20, b=20, l=20, r=20)
    )

    return fig


# =========================================================
# CONTROLE DE FILTRO
# =========================================================
def ensure_state_list(key, options):
    if key not in st.session_state:
        st.session_state[key] = list(options)
    else:
        st.session_state[key] = [x for x in st.session_state[key] if x in options]


def multiselect_with_actions(label, options, key, location="sidebar"):
    options = list(options)
    ensure_state_list(key, options)

    parent = st.sidebar if location == "sidebar" else st

    c1, c2 = parent.columns(2)
    if c1.button("Selecionar tudo", key=f"{key}_all"):
        st.session_state[key] = list(options)
        st.rerun()

    if c2.button("Limpar tudo", key=f"{key}_clear"):
        st.session_state[key] = []
        st.rerun()

    st.session_state[key] = [x for x in st.session_state[key] if x in options]

    selected = parent.multiselect(
        label,
        options=options,
        key=key
    )
    return selected


# =========================================================
# LEITURA / PREPARO
# =========================================================
@st.cache_data(show_spinner=False)
def carregar_excel(uploaded_file):
    return pd.read_excel(uploaded_file, usecols=lambda c: c in COLUNAS_LEITURA)


@st.cache_data(show_spinner=False)
def preparar_base(df_raw):
    df = df_raw.copy()

    if "Cidade destinatária" not in df.columns:
        df["Cidade destinatária"] = np.nan

    for col in ["Data de criação", "Horário de saída para entrega", "Horário da entrega"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    df["base_padronizada"] = df["Base de entrega"].apply(padronizar_base)
    df["entregador_padronizado"] = df["Responsável pela entrega"].apply(padronizar_entregador)
    df["bairro_padronizado"] = df["Distrito destinatário"].apply(padronizar_bairro)
    df["cidade_padronizada"] = df["Cidade destinatária"].apply(padronizar_cidade)
    df["status_sla"] = df["Marca de assinatura"].apply(classificar_status_sla)
    df["motivo_tratado"] = df["Motivos dos pacotes problemáticos"].apply(tratar_motivo)

    df["data_primeira_saida"] = pd.to_datetime(df["Data de criação"], errors="coerce").dt.normalize()
    df["data_saida_atual"] = pd.to_datetime(df["Horário de saída para entrega"], errors="coerce").dt.normalize()
    df["data_entrega"] = pd.to_datetime(df["Horário da entrega"], errors="coerce").dt.normalize()

    # Hora da baixa
    df["hora_entrega"] = pd.to_datetime(df["Horário da entrega"], errors="coerce").dt.hour
    df["hora_entrega_label"] = df["hora_entrega"].apply(
        lambda x: f"{int(x):02d}:00" if pd.notna(x) else np.nan
    )

    df["flag_entregue"] = (df["status_sla"] == "Entregue").astype(int)
    df["flag_nao_entregue"] = (df["status_sla"] == "Não entregue").astype(int)
    df["flag_devolucao"] = (df["status_sla"] == "Devolução").astype(int)

    df["flag_bairro_nao_informado"] = (df["bairro_padronizado"] == "Não informado").astype(int)
    df["flag_cidade_nao_informada"] = (df["cidade_padronizada"] == "Não informado").astype(int)
    df["flag_entregador_nao_informado"] = (df["entregador_padronizado"] == "Não informado").astype(int)
    df["flag_motivo_nao_informado"] = (df["motivo_tratado"] == "Sem motivo informado").astype(int)
    df["flag_data_criacao_invalida"] = df["data_primeira_saida"].isna().astype(int)

    return df


# =========================================================
# BASES LÓGICAS
# =========================================================
def criar_bases_logicas_periodo(df, data_inicio, data_fim):
    data_inicio = pd.Timestamp(data_inicio).normalize()
    data_fim = pd.Timestamp(data_fim).normalize()

    df_sla_periodo = df[
        df["data_primeira_saida"].between(data_inicio, data_fim, inclusive="both")
    ].copy()

    df_pendencia = df[
        (df["status_sla"] == "Não entregue") &
        (df["data_primeira_saida"].notna()) &
        (df["data_primeira_saida"] < data_inicio)
    ].copy()

    if not df_pendencia.empty:
        df_pendencia["dias_pendente"] = (data_inicio - df_pendencia["data_primeira_saida"]).dt.days
        df_pendencia["dias_pendente"] = df_pendencia["dias_pendente"].apply(
            lambda x: max(x, 0) if pd.notna(x) else np.nan
        )
        df_pendencia["faixa_pendente"] = df_pendencia["dias_pendente"].apply(classificar_faixa_pendente)
    else:
        df_pendencia["dias_pendente"] = pd.Series(dtype="float")
        df_pendencia["faixa_pendente"] = pd.Series(dtype="object")

    return df_sla_periodo, df_pendencia


# =========================================================
# CÁLCULOS
# =========================================================
def calcular_indicadores_sla(df):
    total = len(df)
    entregues = int(df["flag_entregue"].sum())
    nao_entregues = int(df["flag_nao_entregue"].sum())
    devolucoes = int(df["flag_devolucao"].sum())
    sla_geral = (entregues / total * 100) if total > 0 else 0

    return {
        "total": total,
        "entregues": entregues,
        "nao_entregues": nao_entregues,
        "devolucoes": devolucoes,
        "sla_geral": sla_geral,
    }


def calcular_indicadores_pendencia(df):
    total = len(df)
    d1 = int((df["dias_pendente"] == 1).sum()) if not df.empty else 0
    d2 = int((df["dias_pendente"] == 2).sum()) if not df.empty else 0
    d3 = int((df["dias_pendente"] >= 3).sum()) if not df.empty else 0
    return {
        "total_pendencia": total,
        "pend_d1": d1,
        "pend_d2": d2,
        "pend_d3": d3,
    }


def calcular_alertas_qualidade(df):
    total = len(df)
    nao_entregues = max(int(df["flag_nao_entregue"].sum()), 1)

    bairro_nao_info = int(df["flag_bairro_nao_informado"].sum())
    cidade_nao_info = int(df["flag_cidade_nao_informada"].sum())
    entregador_nao_info = int(df["flag_entregador_nao_informado"].sum())
    data_invalida = int(df["flag_data_criacao_invalida"].sum())

    base_nao_entregue = df[df["status_sla"] == "Não entregue"].copy()
    motivo_nao_info = int((base_nao_entregue["motivo_tratado"] == "Sem motivo informado").sum())

    return pd.DataFrame([
        {
            "Indicador": "Bairro não informado",
            "Quantidade": bairro_nao_info,
            "% Base": formatar_percentual((bairro_nao_info / total * 100) if total else 0)
        },
        {
            "Indicador": "Cidade não informada",
            "Quantidade": cidade_nao_info,
            "% Base": formatar_percentual((cidade_nao_info / total * 100) if total else 0)
        },
        {
            "Indicador": "Entregador não informado",
            "Quantidade": entregador_nao_info,
            "% Base": formatar_percentual((entregador_nao_info / total * 100) if total else 0)
        },
        {
            "Indicador": "Data de criação inválida",
            "Quantidade": data_invalida,
            "% Base": formatar_percentual((data_invalida / total * 100) if total else 0)
        },
        {
            "Indicador": "Motivo não informado (não entregues)",
            "Quantidade": motivo_nao_info,
            "% Não entregues": formatar_percentual((motivo_nao_info / nao_entregues * 100) if nao_entregues else 0)
        },
    ])


def agregar_sla(df, coluna_grupo):
    if df.empty:
        return pd.DataFrame(columns=[coluna_grupo, "total_pedidos", "entregues", "nao_entregues", "devolucoes", "sla_num", "participacao_volume_num"])

    resumo = (
        df.groupby(coluna_grupo, dropna=False)
        .agg(
            total_pedidos=("Número de pedido JMS", "count"),
            entregues=("flag_entregue", "sum"),
            nao_entregues=("flag_nao_entregue", "sum"),
            devolucoes=("flag_devolucao", "sum"),
        )
        .reset_index()
    )

    resumo["sla_num"] = np.where(
        resumo["total_pedidos"] > 0,
        resumo["entregues"] / resumo["total_pedidos"] * 100,
        0
    )

    resumo["participacao_volume_num"] = np.where(
        resumo["total_pedidos"].sum() > 0,
        resumo["total_pedidos"] / resumo["total_pedidos"].sum() * 100,
        0
    )

    return resumo


def montar_tabela_sla(df, coluna_grupo, nome_coluna_saida, volume_min=1):
    resumo = agregar_sla(df, coluna_grupo)
    if resumo.empty:
        return pd.DataFrame(columns=[
            nome_coluna_saida, "total_pedidos", "entregues", "nao_entregues",
            "devolucoes", "sla_%", "participacao_volume_%"
        ])

    resumo = resumo[resumo["total_pedidos"] >= volume_min].copy()
    resumo = resumo.sort_values(by=["sla_num", "total_pedidos"], ascending=[True, False]).reset_index(drop=True)
    resumo["sla_%"] = resumo["sla_num"].map(formatar_percentual)
    resumo["participacao_volume_%"] = resumo["participacao_volume_num"].map(formatar_percentual)
    resumo = resumo.rename(columns={coluna_grupo: nome_coluna_saida})

    return resumo[
        [nome_coluna_saida, "total_pedidos", "entregues", "nao_entregues", "devolucoes", "sla_%", "participacao_volume_%"]
    ]


def calcular_motivos_nao_entregues(df, top_n=30):
    if df.empty:
        return pd.DataFrame(columns=["motivo_tratado", "qtd"])

    base = df[df["status_sla"] == "Não entregue"].copy()
    resumo = (
        base.groupby("motivo_tratado", dropna=False)
        .agg(qtd=("Número de pedido JMS", "count"))
        .reset_index()
        .sort_values(by="qtd", ascending=False)
        .head(top_n)
        .reset_index(drop=True)
    )
    return resumo


def calcular_pendencias(df):
    if df.empty:
        return pd.DataFrame(columns=["faixa_pendente", "qtd"])

    resumo = (
        df.groupby("faixa_pendente", dropna=False)
        .agg(qtd=("Número de pedido JMS", "count"))
        .reset_index()
    )

    ordem = {"D1": 1, "D2": 2, "D3+": 3}
    resumo["ordem"] = resumo["faixa_pendente"].map(ordem)
    resumo = resumo.sort_values("ordem").drop(columns="ordem")
    return resumo


def obter_diagnostico_padronizacao(df, original, padronizado, limite=200):
    if original not in df.columns or padronizado not in df.columns:
        return pd.DataFrame()

    return (
        df[[original, padronizado]]
        .drop_duplicates()
        .sort_values(by=[padronizado, original])
        .head(limite)
    )


# =========================================================
# PRODUTIVIDADE HORÁRIA
# =========================================================
def calcular_baixas_por_hora(df):
    base = df[
        (df["status_sla"] == "Entregue") &
        (df["hora_entrega"].notna())
    ].copy()

    if base.empty:
        return pd.DataFrame(columns=["hora_entrega", "hora", "qtd_baixas"])

    resumo = (
        base.groupby(["hora_entrega", "hora_entrega_label"], dropna=False)
        .agg(qtd_baixas=("Número de pedido JMS", "count"))
        .reset_index()
        .rename(columns={"hora_entrega_label": "hora"})
        .sort_values(by="hora_entrega")
    )

    return resumo[["hora_entrega", "hora", "qtd_baixas"]]


def montar_matriz_baixas_entregador(df, top_n=15):
    base = df[
        (df["status_sla"] == "Entregue") &
        (df["hora_entrega_label"].notna())
    ].copy()

    if base.empty:
        return pd.DataFrame()

    top_entregadores = (
        base.groupby("entregador_padronizado")
        .agg(total_baixas=("Número de pedido JMS", "count"))
        .reset_index()
        .sort_values(by="total_baixas", ascending=False)
        .head(top_n)["entregador_padronizado"]
        .tolist()
    )

    base = base[base["entregador_padronizado"].isin(top_entregadores)].copy()

    matriz = pd.pivot_table(
        base,
        index="entregador_padronizado",
        columns="hora_entrega_label",
        values="Número de pedido JMS",
        aggfunc="count",
        fill_value=0
    )

    if matriz.empty:
        return pd.DataFrame()

    ordem_colunas = sorted(
        matriz.columns,
        key=lambda x: int(str(x).split(":")[0])
    )
    matriz = matriz[ordem_colunas].reset_index()

    return matriz


def calcular_ranking_baixas_entregador(df, top_n=20):
    base = df[df["status_sla"] == "Entregue"].copy()

    if base.empty:
        return pd.DataFrame(columns=["Entregador", "Baixas", "Participação %"])

    resumo = (
        base.groupby("entregador_padronizado")
        .agg(Baixas=("Número de pedido JMS", "count"))
        .reset_index()
        .rename(columns={"entregador_padronizado": "Entregador"})
        .sort_values(by="Baixas", ascending=False)
        .head(top_n)
        .reset_index(drop=True)
    )

    total = resumo["Baixas"].sum()
    resumo["Participação %"] = resumo["Baixas"].apply(
        lambda x: formatar_percentual((x / total * 100) if total else 0)
    )

    return resumo


# =========================================================
# FAROL POR BASE
# =========================================================
def calcular_farol_por_base(df_sla_periodo, meta_periodo, volume_min=1):
    if df_sla_periodo.empty:
        return pd.DataFrame(columns=["Base", "Total", "Entregues", "Não entregues", "Devoluções", "SLA %", "Meta %", "Gap p.p.", "Farol"])

    resumo = (
        df_sla_periodo.groupby("base_padronizada", dropna=False)
        .agg(
            Total=("Número de pedido JMS", "count"),
            Entregues=("flag_entregue", "sum"),
            NaoEntregues=("flag_nao_entregue", "sum"),
            Devolucoes=("flag_devolucao", "sum"),
        )
        .reset_index()
        .rename(columns={"base_padronizada": "Base"})
    )

    resumo = resumo[resumo["Total"] >= volume_min].copy()
    resumo["SLA_num"] = np.where(resumo["Total"] > 0, resumo["Entregues"] / resumo["Total"] * 100, 0)
    resumo["Meta_num"] = meta_periodo
    resumo["Gap_num"] = resumo["SLA_num"] - resumo["Meta_num"]
    resumo["Farol"] = resumo["Gap_num"].apply(
        lambda x: "🟢 Acima da meta" if x >= 0 else ("🟡 Próximo da meta" if x >= -2 else "🔴 Abaixo da meta")
    )

    resumo["SLA %"] = resumo["SLA_num"].map(formatar_percentual)
    resumo["Meta %"] = resumo["Meta_num"].map(formatar_percentual)
    resumo["Gap p.p."] = resumo["Gap_num"].map(lambda x: f"{x:+.2f}")

    resumo = resumo.rename(columns={
        "NaoEntregues": "Não entregues",
        "Devolucoes": "Devoluções"
    })

    return resumo[
        ["Base", "Total", "Entregues", "Não entregues", "Devoluções", "SLA %", "Meta %", "Gap p.p.", "Farol"]
    ].sort_values(by=["Farol", "Total"], ascending=[True, False]).reset_index(drop=True)


# =========================================================
# INTERFACE
# =========================================================
st.title("Indicador de SLA")
st.caption("Versão V7 com meta dinâmica, rosca do SLA, curva acumulada, heatmap e farol por base.")

arquivo = st.sidebar.file_uploader("Faça upload da base Excel", type=["xlsx", "xls"])

st.sidebar.markdown("---")
st.sidebar.subheader("Parâmetros")

volume_min = st.sidebar.number_input(
    "Volume mínimo para rankings",
    min_value=1,
    max_value=1000,
    value=10,
    step=1
)

limite_detalhe = st.sidebar.number_input(
    "Máximo de linhas no detalhamento",
    min_value=100,
    max_value=50000,
    value=5000,
    step=100
)

if arquivo:
    try:
        with st.spinner("Lendo base..."):
            df_raw = carregar_excel(arquivo)

        faltantes = validar_colunas(df_raw)
        if faltantes:
            st.error("A base não contém todas as colunas obrigatórias.")
            st.write("Colunas faltantes:")
            st.write(faltantes)
            st.stop()

        with st.spinner("Tratando base..."):
            df = preparar_base(df_raw)

        datas_validas = sorted([pd.Timestamp(d).date() for d in df["data_primeira_saida"].dropna().unique()])

        if not datas_validas:
            st.error("Nenhuma data válida foi encontrada na coluna 'Data de criação'.")
            st.stop()

        data_min = min(datas_validas)
        data_max = max(datas_validas)

        st.sidebar.markdown("---")
        st.sidebar.subheader("Período de análise")

        intervalo = st.sidebar.date_input(
            "Selecione o intervalo",
            value=(data_max, data_max),
            min_value=data_min,
            max_value=data_max,
            format="DD/MM/YYYY"
        )

        if isinstance(intervalo, tuple) and len(intervalo) == 2:
            data_inicio, data_fim = intervalo
        else:
            data_inicio = intervalo
            data_fim = intervalo

        if data_inicio > data_fim:
            st.error("A data inicial não pode ser maior que a data final.")
            st.stop()

        st.sidebar.markdown("---")
        st.sidebar.subheader("Filtros")

        bases_opts = sorted(df["base_padronizada"].dropna().unique().tolist())
        bases_sel = multiselect_with_actions("Base", bases_opts, "filtro_base")

        df_base = df[df["base_padronizada"].isin(bases_sel)].copy() if bases_sel else df.iloc[0:0].copy()

        cidades_opts = sorted(df_base["cidade_padronizada"].dropna().unique().tolist())
        cidades_sel = multiselect_with_actions("Cidade", cidades_opts, "filtro_cidade")

        df_cidade = df_base[df_base["cidade_padronizada"].isin(cidades_sel)].copy() if cidades_sel else df_base.iloc[0:0].copy()

        bairros_opts = sorted(df_cidade["bairro_padronizado"].dropna().unique().tolist())
        bairros_sel = multiselect_with_actions("Bairro", bairros_opts, "filtro_bairro")

        df_bairro = df_cidade[df_cidade["bairro_padronizado"].isin(bairros_sel)].copy() if bairros_sel else df_cidade.iloc[0:0].copy()

        entregadores_opts = sorted(df_bairro["entregador_padronizado"].dropna().unique().tolist())
        entregadores_sel = multiselect_with_actions("Entregador", entregadores_opts, "filtro_entregador")

        status_opts = sorted(df["status_sla"].dropna().unique().tolist())
        status_sel = multiselect_with_actions("Status", status_opts, "filtro_status")

        df_filtrado = df[
            df["base_padronizada"].isin(bases_sel)
            & df["cidade_padronizada"].isin(cidades_sel)
            & df["bairro_padronizado"].isin(bairros_sel)
            & df["entregador_padronizado"].isin(entregadores_sel)
            & df["status_sla"].isin(status_sel)
        ].copy()

        df_sla_periodo, df_pendencia = criar_bases_logicas_periodo(df_filtrado, data_inicio, data_fim)

        indicadores_sla = calcular_indicadores_sla(df_sla_periodo)
        indicadores_pend = calcular_indicadores_pendencia(df_pendencia)
        alertas_qualidade = calcular_alertas_qualidade(df_filtrado)

        meta_periodo = calcular_meta_periodo(data_inicio, data_fim)
        gap_meta = indicadores_sla["sla_geral"] - meta_periodo
        farol_periodo = classificar_farol_sla(indicadores_sla["sla_geral"], meta_periodo)

        info1, info2, info3, info4 = st.columns(4)
        with info1:
            st.info(f"SLA do período: {pd.Timestamp(data_inicio).strftime('%d/%m/%Y')} até {pd.Timestamp(data_fim).strftime('%d/%m/%Y')}")
        with info2:
            st.info(f"Pendência acumulada: anteriores a {pd.Timestamp(data_inicio).strftime('%d/%m/%Y')}")
        with info3:
            st.info(f"Meta do período: {meta_periodo:.2f}%")
        with info4:
            st.info(f"Farol do período: {farol_periodo}")

        top_left, top_right = st.columns([2, 1])

        with top_left:
            with st.container(border=True):
                st.subheader("SLA do período")
                c1, c2, c3, c4, c5 = st.columns(5)
                c1.metric("Pacotes do período", formatar_inteiro(indicadores_sla["total"]))
                c2.metric("Entregues", formatar_inteiro(indicadores_sla["entregues"]))
                c3.metric("Não entregues", formatar_inteiro(indicadores_sla["nao_entregues"]))
                c4.metric("Devoluções", formatar_inteiro(indicadores_sla["devolucoes"]))
                c5.metric("SLA do período", formatar_percentual(indicadores_sla["sla_geral"]))

        with top_right:
            with st.container(border=True):
                st.subheader("SLA x Meta")
                fig_rosca = criar_grafico_rosca_sla(indicadores_sla["sla_geral"], meta_periodo)
                st.plotly_chart(fig_rosca, use_container_width=True)

                if gap_meta >= 0:
                    st.success(f"Acima da meta em {gap_meta:.2f} p.p.")
                else:
                    st.error(f"Abaixo da meta em {abs(gap_meta):.2f} p.p.")

        with st.container(border=True):
            st.subheader("Pendência acumulada")
            c6, c7, c8, c9 = st.columns(4)
            c6.metric("Pendentes acumulados", formatar_inteiro(indicadores_pend["total_pendencia"]))
            c7.metric("Pendentes D1", formatar_inteiro(indicadores_pend["pend_d1"]))
            c8.metric("Pendentes D2", formatar_inteiro(indicadores_pend["pend_d2"]))
            c9.metric("Pendentes D3+", formatar_inteiro(indicadores_pend["pend_d3"]))

        tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9, tab10 = st.tabs([
            "Visão Geral",
            "SLA por Base",
            "SLA por Entregador",
            "SLA por Bairro",
            "Pendências",
            "Produtividade Horária",
            "Farol por Base",
            "Qualidade",
            "Detalhamento",
            "Exportação"
        ])

        with tab1:
            a, b = st.columns(2)

            with a:
                with st.container(border=True):
                    st.markdown("**Top bases por volume no período**")
                    tabela_base = montar_tabela_sla(df_sla_periodo, "base_padronizada", "Base", volume_min=1)
                    st.dataframe(
                        tabela_base.sort_values(by="total_pedidos", ascending=False).head(20),
                        use_container_width=True,
                        hide_index=True
                    )

            with b:
                with st.container(border=True):
                    st.markdown("**Motivos de não entrega no período**")
                    st.dataframe(
                        calcular_motivos_nao_entregues(df_sla_periodo, top_n=20),
                        use_container_width=True,
                        hide_index=True
                    )

        with tab2:
            with st.container(border=True):
                st.markdown("**SLA por base**")
                st.dataframe(
                    montar_tabela_sla(df_sla_periodo, "base_padronizada", "Base", volume_min=volume_min),
                    use_container_width=True,
                    hide_index=True
                )

        with tab3:
            with st.container(border=True):
                st.markdown("**SLA por entregador**")
                st.dataframe(
                    montar_tabela_sla(df_sla_periodo, "entregador_padronizado", "Entregador", volume_min=volume_min),
                    use_container_width=True,
                    hide_index=True
                )

        with tab4:
            with st.container(border=True):
                st.markdown("**SLA por bairro**")
                st.dataframe(
                    montar_tabela_sla(df_sla_periodo, "bairro_padronizado", "Bairro", volume_min=volume_min),
                    use_container_width=True,
                    hide_index=True
                )

        with tab5:
            a, b = st.columns([1, 2])

            with a:
                with st.container(border=True):
                    st.markdown("**Faixas de pendência**")
                    st.dataframe(
                        calcular_pendencias(df_pendencia),
                        use_container_width=True,
                        hide_index=True
                    )

            with b:
                with st.container(border=True):
                    st.markdown("**Pedidos não entregues acumulados**")
                    cols_pend = [
                        "Número de pedido JMS",
                        "base_padronizada",
                        "cidade_padronizada",
                        "bairro_padronizado",
                        "entregador_padronizado",
                        "data_primeira_saida",
                        "dias_pendente",
                        "faixa_pendente",
                        "motivo_tratado",
                    ]
                    st.dataframe(
                        df_pendencia[cols_pend]
                        .sort_values(by=["dias_pendente", "base_padronizada", "bairro_padronizado"], ascending=[False, True, True])
                        .head(limite_detalhe),
                        use_container_width=True,
                        hide_index=True
                    )

        with tab6:
            baixas_hora = calcular_baixas_por_hora(df_sla_periodo)

            with st.container(border=True):
                st.markdown("**Baixas por hora — visão geral**")
                if baixas_hora.empty:
                    st.info("Não há entregas com horário de baixa válido no período selecionado.")
                else:
                    grafico_hora = baixas_hora.set_index("hora")[["qtd_baixas"]]
                    st.bar_chart(grafico_hora, use_container_width=True)

                    col_g1, col_g2 = st.columns(2)

                    with col_g1:
                        st.markdown("**Tabela de baixas por hora**")
                        st.dataframe(
                            baixas_hora[["hora", "qtd_baixas"]],
                            use_container_width=True,
                            hide_index=True
                        )

                    with col_g2:
                        st.markdown("**Curva acumulada de baixas**")
                        fig_acumulada = criar_grafico_linha_acumulada(baixas_hora)
                        st.plotly_chart(fig_acumulada, use_container_width=True)

            with st.container(border=True):
                st.markdown("**Baixas por hora por entregador**")

                top_n_entregadores = st.slider(
                    "Quantidade de entregadores na análise",
                    min_value=5,
                    max_value=50,
                    value=15,
                    step=5,
                    key="slider_top_entregadores_hora"
                )

                ranking_entregadores = calcular_ranking_baixas_entregador(
                    df_sla_periodo,
                    top_n=top_n_entregadores
                )

                matriz_entregador = montar_matriz_baixas_entregador(
                    df_sla_periodo,
                    top_n=top_n_entregadores
                )

                if ranking_entregadores.empty or matriz_entregador.empty:
                    st.info("Não há entregas suficientes para análise por entregador.")
                else:
                    c1, c2 = st.columns([1, 2])

                    with c1:
                        st.markdown("**Ranking de baixas por entregador**")
                        st.dataframe(
                            ranking_entregadores,
                            use_container_width=True,
                            hide_index=True
                        )

                    with c2:
                        st.markdown("**Matriz de baixas por hora x entregador**")
                        st.dataframe(
                            matriz_entregador,
                            use_container_width=True,
                            hide_index=True
                        )

                    st.markdown("**Heatmap de baixas por entregador x hora**")
                    fig_heatmap = criar_heatmap_entregador_hora(matriz_entregador)
                    st.plotly_chart(fig_heatmap, use_container_width=True)

        with tab7:
            with st.container(border=True):
                st.markdown("**Farol por base**")
                tabela_farol = calcular_farol_por_base(
                    df_sla_periodo,
                    meta_periodo=meta_periodo,
                    volume_min=volume_min
                )
                st.dataframe(
                    tabela_farol,
                    use_container_width=True,
                    hide_index=True
                )

        with tab8:
            with st.container(border=True):
                st.markdown("**Qualidade da base filtrada**")
                st.dataframe(alertas_qualidade, use_container_width=True, hide_index=True)

            with st.expander("Diagnóstico de padronização de bairros"):
                st.dataframe(
                    obter_diagnostico_padronizacao(df_filtrado, "Distrito destinatário", "bairro_padronizado", limite=500),
                    use_container_width=True,
                    hide_index=True
                )

            with st.expander("Diagnóstico de padronização de cidades"):
                st.dataframe(
                    obter_diagnostico_padronizacao(df_filtrado, "Cidade destinatária", "cidade_padronizada", limite=500),
                    use_container_width=True,
                    hide_index=True
                )

            with st.expander("Diagnóstico de padronização de entregadores"):
                st.dataframe(
                    obter_diagnostico_padronizacao(df_filtrado, "Responsável pela entrega", "entregador_padronizado", limite=500),
                    use_container_width=True,
                    hide_index=True
                )

            with st.expander("Textos originais de status"):
                diagnostico_status = (
                    df_filtrado[["Marca de assinatura", "status_sla"]]
                    .drop_duplicates()
                    .sort_values(by=["status_sla", "Marca de assinatura"])
                )
                st.dataframe(diagnostico_status, use_container_width=True, hide_index=True)

        with tab9:
            tipo_detalhamento = st.radio(
                "Escolha o conjunto",
                ["SLA do período", "Pendência acumulada", "Base filtrada completa"],
                horizontal=True
            )

            if tipo_detalhamento == "SLA do período":
                df_det = df_sla_periodo.copy()
            elif tipo_detalhamento == "Pendência acumulada":
                df_det = df_pendencia.copy()
            else:
                df_det = df_filtrado.copy()

            colunas_exibir = [
                "Número de pedido JMS",
                "base_padronizada",
                "cidade_padronizada",
                "bairro_padronizado",
                "entregador_padronizado",
                "status_sla",
                "motivo_tratado",
                "Data de criação",
                "Horário de saída para entrega",
                "Horário da entrega",
                "hora_entrega_label",
            ]

            if "dias_pendente" in df_det.columns:
                colunas_exibir += ["dias_pendente", "faixa_pendente"]

            with st.container(border=True):
                st.dataframe(
                    df_det[colunas_exibir].head(limite_detalhe),
                    use_container_width=True,
                    hide_index=True
                )

        with tab10:
            tabela_base_exp = montar_tabela_sla(df_sla_periodo, "base_padronizada", "Base", volume_min=volume_min)
            tabela_entregador_exp = montar_tabela_sla(df_sla_periodo, "entregador_padronizado", "Entregador", volume_min=volume_min)
            tabela_bairro_exp = montar_tabela_sla(df_sla_periodo, "bairro_padronizado", "Bairro", volume_min=volume_min)
            pendencias_exp = calcular_pendencias(df_pendencia)
            baixas_hora_exp = calcular_baixas_por_hora(df_sla_periodo)
            matriz_hora_entregador_exp = montar_matriz_baixas_entregador(df_sla_periodo, top_n=30)
            ranking_baixas_exp = calcular_ranking_baixas_entregador(df_sla_periodo, top_n=30)
            farol_base_exp = calcular_farol_por_base(df_sla_periodo, meta_periodo=meta_periodo, volume_min=volume_min)

            excel_bytes = dataframe_to_excel_bytes({
                "SLA_Base": tabela_base_exp,
                "SLA_Entregador": tabela_entregador_exp,
                "SLA_Bairro": tabela_bairro_exp,
                "Pendencias": pendencias_exp,
                "Baixas_por_Hora": baixas_hora_exp,
                "Baixas_Hora_Entregador": matriz_hora_entregador_exp,
                "Ranking_Baixas": ranking_baixas_exp,
                "Farol_Base": farol_base_exp,
                "Detalhe_SLA_Periodo": df_sla_periodo.head(limite_detalhe),
                "Detalhe_Pendencia": df_pendencia.head(limite_detalhe),
            })

            with st.container(border=True):
                st.download_button(
                    label="Baixar resultados em Excel",
                    data=excel_bytes,
                    file_name=f"indicador_sla_{pd.Timestamp(data_inicio).strftime('%Y%m%d')}_{pd.Timestamp(data_fim).strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        st.caption(
            f"Linhas carregadas: {formatar_inteiro(len(df_raw))} | "
            f"Após filtros: {formatar_inteiro(len(df_filtrado))} | "
            f"SLA do período: {formatar_inteiro(len(df_sla_periodo))} | "
            f"Pendência acumulada: {formatar_inteiro(len(df_pendencia))}"
        )

    except Exception as e:
        st.error("Erro ao processar a base.")
        st.exception(e)

else:
    st.info("Faça o upload da base Excel para iniciar a análise.")
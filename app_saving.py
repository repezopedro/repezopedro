import streamlit as st
import pandas as pd
import plotly.express as px
import os
import json

# --------------------------------------------------------------
# Caminho fixo do arquivo Excel
# --------------------------------------------------------------
CAMINHO_ARQUIVO = r"C:\Users\pedro.muniz\Documents\pythonlocal\BIOXXI\dashboard_compras\saving\compras_2025_colig1.xlsx"

# --------------------------------------------------------------

# Caminho do arquivo para salvar os fatores de conversão
CAMINHO_FATOR_CONVERSAO = os.path.join(os.path.dirname(__file__), "fator_conversao.json")

# Carregar fatores de conversão salvos, se existirem
def carregar_fatores():
    if os.path.exists(CAMINHO_FATOR_CONVERSAO):
        with open(CAMINHO_FATOR_CONVERSAO, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

# Salvar fatores de conversão
def salvar_fatores(fatores):
    with open(CAMINHO_FATOR_CONVERSAO, "w", encoding="utf-8") as f:
        json.dump(fatores, f, ensure_ascii=False, indent=2)

FATOR_CONVERSAO = carregar_fatores()


# Função para converter QTDENTRADA para unidade padrão (UN)
def converter_para_unidade(row):
    fator = FATOR_CONVERSAO.get(str(row["CODUND"]).upper(), 1)
    return row["QTDENTRADA"] * fator


# --------------------------------------------------------------
# Função: Calcular saving e preparar os dados
# --------------------------------------------------------------
def processar_dados(compras, negociacao):
    compras.columns = compras.columns.str.strip().str.upper()
    negociacao.columns = negociacao.columns.str.strip().str.upper()

    colunas_necessarias_compras = [
        "NUMERO_MOVIMENTO", "CODIGO_PRD", "CODCOLIGADA", "NOME_RAZAO_SOCIAL_FORNECEDOR",
        "PRECODENTRADA", "QTDENTRADA", "DATA_ENTRADA", "NOME_PRODUTO", "CODUND"
    ]
    colunas_necessarias_negociacao = [
        "CODCOLIGADA", "CODIGO_PRD", "PRECO_ANTIGO", "NOME_RAZAO_SOCIAL_FORNECEDOR", "FORNECEDOR_ANTIGO"
    ]

    for c in colunas_necessarias_compras:
        if c not in compras.columns:
            st.warning(f"⚠️ Coluna '{c}' não encontrada na aba de COMPRAS.")
    for c in colunas_necessarias_negociacao:
        if c not in negociacao.columns:
            st.warning(f"⚠️ Coluna '{c}' não encontrada na aba de NEGOCIAÇÃO.")

    # Merge para identificar onde o fornecedor da negociação aparece na tabela compras
    dados = pd.merge(
        compras,
        negociacao,
        how="left",
        left_on=["CODIGO_PRD", "CODCOLIGADA", "NOME_RAZAO_SOCIAL_FORNECEDOR"],
        right_on=["CODIGO_PRD", "CODCOLIGADA", "NOME_RAZAO_SOCIAL_FORNECEDOR"],
        suffixes=("_COMPRA", "_NEGOCIACAO")
    )

    # Adiciona coluna QTDENTRADA_UN convertida
    dados["QTDENTRADA_UN"] = dados.apply(converter_para_unidade, axis=1)

    # 🔍 DEBUG: Mostrar fornecedores e correspondências
    #st.write("🧩 Fornecedores únicos em COMPRAS:", compras["NOME_RAZAO_SOCIAL_FORNECEDOR"].unique())
    #st.write("🧩 Fornecedores únicos em NEGOCIAÇÃO:", negociacao["NOME_RAZAO_SOCIAL_FORNECEDOR"].unique())


    # Calcula saving apenas para IDs em que o fornecedor da negociação aparece
    mask_novo_forn = dados["PRECO_ANTIGO"].notna() & dados["PRECODENTRADA"].notna() & dados["QTDENTRADA_UN"].notna()
    dados.loc[mask_novo_forn, "SAVING"] = (
        (dados.loc[mask_novo_forn, "PRECO_ANTIGO"] * dados.loc[mask_novo_forn, "QTDENTRADA_UN"]) -
        dados.loc[mask_novo_forn, "PRECODENTRADA"]
    )

    # Zera valores negativos
    dados.loc[dados["SAVING"] < 0, "SAVING"] = 0.0
    dados["SAVING_POSITIVO"] = dados["SAVING"] > 0

    # Coluna Mês/Ano
    dados["DATA_ENTRADA"] = pd.to_datetime(dados["DATA_ENTRADA"], errors="coerce")
    dados["MÊS_ANO"] = dados["DATA_ENTRADA"].dt.to_period("M").astype(str)

    return dados

# --------------------------------------------------------------
# Função: Agregar dados mensalmente
# --------------------------------------------------------------
def gerar_resumo(dados):
    resumo = (
        dados.groupby(
            ["MÊS_ANO", "NOME_RAZAO_SOCIAL_FORNECEDOR", "CODIGO_PRD", "NOME_PRODUTO", "CODUND"],
            as_index=False
        )
        .agg({
            "QTDENTRADA": "sum",
            "SAVING": "sum"
        })
        .sort_values(by=["MÊS_ANO", "CODIGO_PRD"])
    )
    resumo["SAVING"] = resumo["SAVING"].round(2)
    return resumo


# --------------------------------------------------------------
# Configuração da página
# --------------------------------------------------------------
st.set_page_config(
    page_title="Painel de Análise de Saving em Compras",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📊 Painel de Análise de Saving em Compras")
st.markdown("""
Este painel mostra **economias reais obtidas quando há troca de fornecedor**.
Somente entradas com **novo fornecedor** são consideradas.
""")

# --------------------------------------------------------------
# Leitura do arquivo Excel
# --------------------------------------------------------------
if not os.path.exists(CAMINHO_ARQUIVO):
    st.error(f"❌ O arquivo não foi encontrado em:\n`{CAMINHO_ARQUIVO}`")
    st.stop()

try:
    sheet_names = pd.ExcelFile(CAMINHO_ARQUIVO).sheet_names
    compras = pd.read_excel(CAMINHO_ARQUIVO, sheet_name=sheet_names[0])
    negociacao = pd.read_excel(CAMINHO_ARQUIVO, sheet_name=sheet_names[1] if len(sheet_names) > 1 else sheet_names[0])

    st.sidebar.header("⚙️ Configurações")
    st.sidebar.markdown(f"**Planilha de Compras:** {sheet_names[0]}")
    st.sidebar.markdown(f"**Planilha de Negociação:** {sheet_names[1] if len(sheet_names) > 1 else sheet_names[0]}")



    # Gerar dicionário FATOR_CONVERSAO dinamicamente, mantendo valores já salvos
    codunds = compras["CODUND"].dropna().unique()
    atualizados = False
    for codun in codunds:
        codun_str = str(codun).upper()
        if codun_str not in FATOR_CONVERSAO:
            FATOR_CONVERSAO[codun_str] = 1
            atualizados = True
    if atualizados:
        salvar_fatores(FATOR_CONVERSAO)


    # Interface para editar fatores manualmente (apenas se clicar no expander)
    with st.sidebar.expander("✏️ Editar Fator de Conversão", expanded=False):
        with st.form("editar_fatores"):
            novos_fatores = {}
            for codun in sorted(FATOR_CONVERSAO.keys()):
                valor = st.number_input(f"{codun}", value=float(FATOR_CONVERSAO[codun]), step=0.01, format="%g")
                novos_fatores[codun] = valor
            salvar = st.form_submit_button("Salvar Fatores")
            if salvar:
                FATOR_CONVERSAO.update(novos_fatores)
                salvar_fatores(FATOR_CONVERSAO)
                st.success("Fatores de conversão salvos!")

    # Processa dados
    dados = processar_dados(compras, negociacao)

    # --------------------------------------------------------------
    # Filtros
    # --------------------------------------------------------------
    colunas_disp = list(dados.columns)
    col_conta = next((c for c in colunas_disp if "CONTA" in c.upper() and ("CONTABIL" in c.upper() or "CONTÁBIL" in c.upper())), None)
    col_natureza = next((c for c in colunas_disp if "NATUR" in c.upper()), None)

    dados_filtrados = dados.copy()

    if col_natureza:
        natureza_opcoes = sorted(dados[col_natureza].dropna().unique().tolist())
        natureza_sel = st.sidebar.multiselect("Filtrar por Natureza:", natureza_opcoes, default=natureza_opcoes)
        dados_filtrados = dados_filtrados[dados_filtrados[col_natureza].isin(natureza_sel)]

    if col_conta:
        conta_opcoes = sorted(dados[col_conta].dropna().astype(str).unique().tolist())
        conta_sel = st.sidebar.multiselect("Filtrar por Conta Contábil:", conta_opcoes, default=conta_opcoes)
        dados_filtrados = dados_filtrados[dados_filtrados[col_conta].astype(str).isin(conta_sel)]

    produtos_opcoes = sorted(dados_filtrados["NOME_PRODUTO"].dropna().unique().tolist())
    produto_sel = st.sidebar.multiselect("Filtrar por Produto:", produtos_opcoes, default=produtos_opcoes)
    dados_filtrados = dados_filtrados[dados_filtrados["NOME_PRODUTO"].isin(produto_sel)]

    # --------------------------------------------------------------
    # Agregação e KPIs
    # --------------------------------------------------------------
    resumo = gerar_resumo(dados_filtrados)

    total_saving = dados_filtrados.loc[dados_filtrados["SAVING_POSITIVO"], "SAVING"].sum()
    total_produtos_saving = dados_filtrados.loc[dados_filtrados["SAVING_POSITIVO"], "CODIGO_PRD"].nunique()
    total_registros = len(dados_filtrados)

    col1, col2, col3 = st.columns(3)
    col1.metric("💰 Economias Alcançadas", f"R$ {total_saving:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    col2.metric("🏷️ Itens Negociados", total_produtos_saving)
    col3.metric("🧾 Registros Filtrados", total_registros)

    # --------------------------------------------------------------
    # Tabela detalhada
    # --------------------------------------------------------------
    st.subheader("📋 Tabela Detalhada (Economias Apuradas)")
    mostrar_somente_saving = st.checkbox("Mostrar apenas > zero", value=True)
    if mostrar_somente_saving:
        resumo_exibir = resumo[resumo["SAVING"] > 0].copy()
    else:
        resumo_exibir = resumo.copy()
    resumo_fmt = resumo_exibir.copy()
    resumo_fmt["SAVING"] = resumo_fmt["SAVING"].apply(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    st.dataframe(resumo_fmt, use_container_width=True)

    # --------------------------------------------------------------
    # Gráfico de barras - evolução mensal
    # --------------------------------------------------------------
    saving_mensal = (
        dados_filtrados[dados_filtrados["SAVING_POSITIVO"]]
        .groupby("MÊS_ANO", as_index=False)["SAVING"].sum()
        .sort_values("MÊS_ANO")
    )

    fig = px.bar(
        saving_mensal,
        x="MÊS_ANO",
        y="SAVING",
        text_auto=".2s",
        title="💹 Evolução Mensal das Economias Alcançadas",
        labels={"SAVING": "Saving Total (R$)", "MÊS_ANO": "Mês/Ano"},
        color="SAVING",
        color_continuous_scale="Greens"
    )
    fig.update_layout(xaxis_tickangle=-45, height=500)
    st.plotly_chart(fig, use_container_width=True)

    # --------------------------------------------------------------
    # Gráfico de linha - saving por produto
    # --------------------------------------------------------------
    st.subheader("📈 Saving por Produto")
    produto_opcoes_grafico = sorted(resumo["NOME_PRODUTO"].unique().tolist())
    produto_sel_grafico = st.selectbox("Selecione um produto:", produto_opcoes_grafico)
    dados_produto = resumo[resumo["NOME_PRODUTO"] == produto_sel_grafico]

    fig2 = px.line(
        dados_produto,
        x="MÊS_ANO",
        y="SAVING",
        markers=True,
        title=f"📊 Evolução do Saving - {produto_sel_grafico}",
        labels={"SAVING": "Saving (R$)", "MÊS_ANO": "Mês/Ano"}
    )
    st.plotly_chart(fig2, use_container_width=True)

    # --------------------------------------------------------------
    # Exportação CSV
    # --------------------------------------------------------------
    st.markdown("### 💾 Exportar Resultados")
    csv = resumo.to_csv(index=False, sep=";", decimal=",").encode("utf-8")
    st.download_button("📥 Baixar resumo filtrado (CSV)", data=csv, file_name="resumo_saving_filtrado.csv", mime="text/csv")

except Exception as e:
    st.error(f"Erro ao processar o arquivo Excel: {e}")

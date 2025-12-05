import streamlit as st
import pandas as pd
import plotly.express as px
import os
import json
import io
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from fpdf import FPDF
from fpdf.enums import Align 

# --------------------------------------------------------------
# Configuração da página (DEVE SER O PRIMEIRO COMANDO STREAMLIT)
# --------------------------------------------------------------
st.set_page_config(
    page_title="Painel de Análise de Saving em Compras",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --------------------------------------------------------------
# Caminhos e Constantes
# --------------------------------------------------------------
CAMINHO_ARQUIVO = r"C:\Users\pedro.muniz\Documents\pythonlocal\BIOXXI\dashboard_compras\saving\compras_2025_colig1.xlsx"
# CAMINHO_ARQUIVO = "compras_2025_colig1.xlsx" # Exemplo para teste local

CAMINHO_FATOR_CONVERSAO = os.path.join(os.path.dirname(__file__), "fator_conversao.json")
CAMINHO_LOGO = os.path.join(os.path.dirname(__file__), "bioxxi_logo.png") 

# --------------------------------------------------------------
# Funções Auxiliares
# --------------------------------------------------------------
def carregar_fatores():
    if os.path.exists(CAMINHO_FATOR_CONVERSAO):
        try:
            with open(CAMINHO_FATOR_CONVERSAO, "r", encoding="utf-8") as f:
                return json.load(f)
        except json.JSONDecodeError:
            return {} 
    return {}

def salvar_fatores(fatores):
    with open(CAMINHO_FATOR_CONVERSAO, "w", encoding="utf-8") as f:
        json.dump(fatores, f, ensure_ascii=False, indent=2)

FATOR_CONVERSAO = carregar_fatores()

def converter_para_unidade(row):
    fator = FATOR_CONVERSAO.get(str(row["CODUND"]).upper(), 1)
    try:
        fator_numerico = float(fator if fator not in [None, ''] else 1)
    except (ValueError, TypeError):
        fator_numerico = 1
    return row["QTDENTRADA"] * fator_numerico

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

    compras["CODIGO_PRD"] = compras["CODIGO_PRD"].astype(str)
    negociacao["CODIGO_PRD"] = negociacao["CODIGO_PRD"].astype(str)
    compras["CODCOLIGADA"] = compras["CODCOLIGADA"].astype(str)
    negociacao["CODCOLIGADA"] = negociacao["CODCOLIGADA"].astype(str)

    dados = pd.merge(
        compras,
        negociacao,
        how="left",
        left_on=["CODIGO_PRD", "CODCOLIGADA", "NOME_RAZAO_SOCIAL_FORNECEDOR"],
        right_on=["CODIGO_PRD", "CODCOLIGADA", "NOME_RAZAO_SOCIAL_FORNECEDOR"],
        suffixes=("_COMPRA", "_NEGOCIACAO")
    )

    dados["QTDENTRADA_UN"] = dados.apply(converter_para_unidade, axis=1)

    mask_novo_forn = dados["PRECO_ANTIGO"].notna() & dados["PRECODENTRADA"].notna() & dados["QTDENTRADA_UN"].notna()
    
    dados["PRECO_ANTIGO"] = pd.to_numeric(dados["PRECO_ANTIGO"], errors='coerce')
    dados["PRECODENTRADA"] = pd.to_numeric(dados["PRECODENTRADA"], errors='coerce')
    
    dados.loc[mask_novo_forn, "SAVING"] = (
        (dados.loc[mask_novo_forn, "PRECO_ANTIGO"] * dados.loc[mask_novo_forn, "QTDENTRADA_UN"]) -
        dados.loc[mask_novo_forn, "PRECODENTRADA"]
    )

    dados.loc[dados["SAVING"] < 0, "SAVING"] = 0.0
    dados["SAVING_POSITIVO"] = dados["SAVING"] > 0
    dados["DATA_ENTRADA"] = pd.to_datetime(dados["DATA_ENTRADA"], errors="coerce")
    dados["MÊS_ANO"] = dados["DATA_ENTRADA"].dt.to_period("M").astype(str)
    dados["CODCOLIGADA"] = pd.to_numeric(dados["CODCOLIGADA"], errors='coerce').fillna(0).astype(int)

    return dados

def gerar_resumo(dados):
    resumo = (
        dados.groupby(
            ["MÊS_ANO", "NOME_RAZAO_SOCIAL_FORNECEDOR", "CODIGO_PRD", "NOME_PRODUTO", "CODUND"],
            as_index=False,
            dropna=False
        )
        .agg({"QTDENTRADA": "sum", "SAVING": "sum"})
        .sort_values(by=["MÊS_ANO", "CODIGO_PRD"])
    )
    resumo["SAVING"] = resumo["SAVING"].round(2)
    return resumo

# --------------------------------------------------------------
# Geração de Gráfico Estático com Matplotlib para o PDF
# --------------------------------------------------------------
def gerar_grafico_fornecedor_img(df):
    """Gera um gráfico de barras horizontais acumulado por fornecedor."""
    df_chart = df.groupby("NOME_RAZAO_SOCIAL_FORNECEDOR")["SAVING"].sum().reset_index()
    df_chart = df_chart[df_chart["SAVING"] > 0].sort_values("SAVING", ascending=True)

    plt.style.use('ggplot')
    fig, ax = plt.subplots(figsize=(10, len(df_chart) * 0.5 + 2)) 
    
    bars = ax.barh(df_chart["NOME_RAZAO_SOCIAL_FORNECEDOR"], df_chart["SAVING"], color='#2E86C1')
    
    ax.set_title("Saving Acumulado por Fornecedor (Ordem Crescente)", fontsize=12, pad=20)
    ax.set_xlabel("Saving Total (R$)", fontsize=10)
    ax.xaxis.set_major_formatter(ticker.StrMethodFormatter('{x:,.0f}'))
    
    for bar in bars:
        width = bar.get_width()
        label_x_pos = width + (width * 0.01)
        ax.text(label_x_pos, bar.get_y() + bar.get_height()/2, 
                f'R$ {width:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.'), 
                va='center', fontsize=8)

    plt.tight_layout()
    
    img_buf = io.BytesIO()
    plt.savefig(img_buf, format='png', bbox_inches='tight', dpi=100)
    img_buf.seek(0)
    plt.close(fig)
    return img_buf

# --------------------------------------------------------------
# Classe PDF Otimizada
# --------------------------------------------------------------
class PDF(FPDF):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.set_doc_option('core_fonts_encoding', 'latin1')
        self.total_saving = 0
        self.total_produtos_saving = 0
        self.total_registros = 0

    def header(self):
        if os.path.exists(CAMINHO_LOGO):
            self.image(CAMINHO_LOGO, 10, 8, 33)
        self.set_font('Arial', 'B', 15)
        self.cell(0, 10, 'RELATÓRIO ANALÍTICO DE SAVING ACUMULADO', ln=True, align='C')
        self.ln(5)
        self.set_draw_color(0, 0, 0)
        self.line(10, 30, self.w - 10, 30)
        self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()}/{{nb}}', 0, 0, 'C')

    def chapter_title(self, title):
        self.set_font('Arial', 'B', 11) 
        self.cell(0, 8, title, ln=True, align='L')
        self.ln(1)

    def chapter_body(self, body):
        self.set_font('Arial', '', 9)
        self.multi_cell(0, 5, body)
        self.ln()

    def add_kpis_and_summary(self):
        self.set_font('Arial', 'B', 10)
        self.cell(0, 6, "Resumo e Indicadores Chave:", ln=True, align='L')
        self.ln(1)
        self.set_font('Arial', '', 9)
        self.cell(0, 5, f"  - Economias Alcançadas: R$ {self.total_saving:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), ln=True)
        self.cell(0, 5, f"  - Itens Negociados: {self.total_produtos_saving}", ln=True)
        self.cell(0, 5, f"  - Registros Filtrados: {self.total_registros}", ln=True)
        self.ln(3)

    def create_table(self, df):
        # --- CONFIGURAÇÃO ---
        self.set_font("Arial", size=6) 
        line_height = 3.5 
        
        df_pdf = df.copy()
        
        # Mapa de nomes
        col_names_map = {
            "MÊS_ANO": "Mês/Ano", "NOME_RAZAO_SOCIAL_FORNECEDOR": "Fornecedor",
            "CODIGO_PRD": "Cód.", "NOME_PRODUTO": "Produto", "CODUND": "UN",
            "QTDENTRADA": "Qtd.", "SAVING": "Saving (R$)",
            "CONSUMO_MEDIO_MENSAL": "Cons. Médio", "SAVING_MEDIO_POR_UNIDADE": "Sav. Méd/UN",
            "PROJECAO_12_MESES": "Proj. 12 Meses"
        }
        df_pdf = df_pdf.rename(columns=col_names_map)
        colunas_pdf = [c for c in df_pdf.columns if c in col_names_map.values()]

        # Formatação de números (com try/except para não quebrar na linha de TOTAL)
        cols_to_format = ["Qtd.", "Saving (R$)", "Cons. Médio", "Sav. Méd/UN", "Proj. 12 Meses"]
        for col in cols_to_format:
            if col in df_pdf.columns:
                def format_value(x):
                    if isinstance(x, (int, float)):
                        return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                    return str(x)
                df_pdf[col] = df_pdf[col].apply(format_value)

        # Larguras
        largura_total = 277.0
        larguras_fixas = {
            "Produto": 80, "Fornecedor": 55, "Proj. 12 Meses": 25, "Saving (R$)": 25,
            "Mês/Ano": 18, "Cód.": 18, "UN": 10, "Qtd.": 15, "Cons. Médio": 20, "Sav. Méd/UN": 20
        }
        larguras_usadas = {k: v for k, v in larguras_fixas.items() if k in colunas_pdf}
        total_estimado = sum(larguras_usadas.values())
        fator_ajuste = largura_total / total_estimado if total_estimado > 0 else 1
        col_widths = [larguras_usadas.get(col, 20) * fator_ajuste for col in colunas_pdf]

        # --- CABEÇALHO ---
        self.set_fill_color(200, 220, 255)
        self.set_font('Arial', 'B', 6)
        x_start = 10 
        self.set_x(x_start)
        for i, col in enumerate(colunas_pdf):
            self.cell(col_widths[i], 5, col, border=1, fill=True, align='C') 
        self.ln(5)

        self.set_font('Arial', '', 6)
        self.set_fill_color(255, 255, 255)

        # --- CORPO DA TABELA (Lógica corrigida para evitar páginas brancas) ---
        for index, row in df_pdf.iterrows():
            # Verifica se é a linha de TOTAL para por negrito
            is_total = str(row[colunas_pdf[0]]).startswith("TOTAL")
            if is_total:
                self.set_font('Arial', 'B', 6)
                self.set_fill_color(240, 240, 240) # Cinza claro para o total
            else:
                self.set_font('Arial', '', 6)
                self.set_fill_color(255, 255, 255)

            # 1. Calcular altura máxima da linha ANTES de desenhar
            max_lines = 1
            data_row = []
            for i, col in enumerate(colunas_pdf):
                texto = str(row[col])
                data_row.append(texto)
                cw = col_widths[i]
                # Estimativa de caracteres por linha
                char_limit = int(cw / 1.6) 
                if len(texto) > char_limit:
                    lines = (len(texto) // char_limit) + 1
                    if lines > max_lines: max_lines = lines
            
            row_height = max_lines * line_height
            
            # 2. Verificar quebra de página
            if self.get_y() + row_height > self.h - 15:
                self.add_page()
                # Reimprime cabeçalho
                self.set_fill_color(200, 220, 255)
                self.set_font('Arial', 'B', 6)
                self.set_x(x_start)
                for i, col in enumerate(colunas_pdf):
                    self.cell(col_widths[i], 5, col, border=1, fill=True, align='C')
                self.ln(5)
                # Restaura fonte da linha
                if is_total:
                    self.set_font('Arial', 'B', 6)
                    self.set_fill_color(240, 240, 240)
                else:
                    self.set_font('Arial', '', 6)
                    self.set_fill_color(255, 255, 255)

            # 3. Desenhar a linha (Agorá é seguro)
            y_curr = self.get_y()
            x_curr = x_start
            
            for i, text in enumerate(data_row):
                w = col_widths[i]
                self.set_xy(x_curr, y_curr)
                self.multi_cell(w, line_height, text, border=0, align='C')
                # Desenha borda da altura total da linha
                self.rect(x_curr, y_curr, w, row_height)
                x_curr += w
            
            # Move Y para a próxima linha
            self.set_y(y_curr + row_height)

def to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Resumo_Saving')
    return output.getvalue()

def to_pdf_bytes(df_resumo, df_projecao, total_saving, total_produtos_saving, total_registros, df_bruto):
    
    # --- 1. PREPARAR TOTAIS (Antes de converter para string sanitizada) ---
    
    # Total Resumo
    if not df_resumo.empty:
        total_qtd = df_resumo["QTDENTRADA"].sum()
        total_saving_val = df_resumo["SAVING"].sum()
        # Cria linha de total
        # A ordem deve bater com as colunas do df_resumo
        # Colunas: MÊS_ANO, NOME_RAZAO_SOCIAL_FORNECEDOR, CODIGO_PRD, NOME_PRODUTO, CODUND, QTDENTRADA, SAVING
        row_total_resumo = {
            "MÊS_ANO": "TOTAL GERAL",
            "NOME_RAZAO_SOCIAL_FORNECEDOR": "",
            "CODIGO_PRD": "",
            "NOME_PRODUTO": "",
            "CODUND": "",
            "QTDENTRADA": total_qtd,
            "SAVING": total_saving_val
        }
        df_resumo = pd.concat([df_resumo, pd.DataFrame([row_total_resumo])], ignore_index=True)

    # Total Projeção
    if not df_projecao.empty:
        total_proj = df_projecao["PROJECAO_12_MESES"].sum()
        # Colunas: NOME_PRODUTO, CONSUMO_MEDIO_MENSAL, SAVING_MEDIO_POR_UNIDADE, PROJECAO_12_MESES
        row_total_proj = {
            "NOME_PRODUTO": "TOTAL GERAL",
            "CONSUMO_MEDIO_MENSAL": "",
            "SAVING_MEDIO_POR_UNIDADE": "",
            "PROJECAO_12_MESES": total_proj
        }
        df_projecao = pd.concat([df_projecao, pd.DataFrame([row_total_proj])], ignore_index=True)

    # --- 2. SANITIZAÇÃO ---
    def sanitize_df(df):
        df_clean = df.copy() 
        for col in df_clean.select_dtypes(include=['object']).columns:
            df_clean[col] = df_clean[col].astype(str).apply(lambda x: x.encode('latin-1', errors='replace').decode('latin-1'))
        return df_clean

    df_resumo_clean = sanitize_df(df_resumo)
    df_projecao_clean = sanitize_df(df_projecao)
    
    # --- 3. GERAÇÃO PDF ---
    pdf = PDF(orientation='L', unit='mm', format='A4')
    pdf.alias_nb_pages()
    pdf.total_saving = total_saving
    pdf.total_produtos_saving = total_produtos_saving
    pdf.total_registros = total_registros

    # Página 1
    pdf.add_page()
    pdf.add_kpis_and_summary()

    pdf.chapter_title("Tabela Detalhada de Economias Apuradas")
    if not df_resumo_clean.empty:
        pdf.create_table(df_resumo_clean)
    else:
        pdf.chapter_body("Nenhum dado de saving apurado.")

    # Gráfico
    if pdf.get_y() > 130: 
        pdf.add_page()
    else:
        pdf.ln(10)

    pdf.chapter_title("Gráfico: Saving Acumulado por Fornecedor (Ordem Crescente)")
    if not df_bruto.empty:
        try:
            img_buffer = gerar_grafico_fornecedor_img(df_bruto)
            pdf.image(img_buffer, w=200, type='PNG') 
            pdf.ln(5)
        except Exception as e:
            pdf.chapter_body(f"Não foi possível gerar o gráfico: {str(e)}")

    # Projeção
    if pdf.get_y() > 160: 
        pdf.add_page()
    else:
        pdf.ln(10)
        
    pdf.chapter_title("Projeção de Saving para 12 Meses por Produto")
    if not df_projecao_clean.empty:
        pdf.create_table(df_projecao_clean)
    else:
        pdf.chapter_body("Nenhuma projeção disponível.")

    saida = pdf.output(dest='S')
    if isinstance(saida, (bytes, bytearray)):
        return bytes(saida)
    else:
        return saida.encode('latin-1')

# --------------------------------------------------------------
# Interface Principal
# --------------------------------------------------------------
st.title("📊 Painel de Análise de Saving em Compras")
st.markdown("""
Este painel mostra economias reais obtidas quando há troca de fornecedor.
Somente entradas com novo fornecedor são consideradas.
""")

if not os.path.exists(CAMINHO_ARQUIVO):
    st.error(f"❌ O arquivo não foi encontrado em:\n`{CAMINHO_ARQUIVO}`")
    st.stop()

# Início do bloco TRY principal
try:
    sheet_names = pd.ExcelFile(CAMINHO_ARQUIVO).sheet_names
    compras = pd.read_excel(CAMINHO_ARQUIVO, sheet_name=sheet_names[0])
    negociacao = pd.read_excel(CAMINHO_ARQUIVO, sheet_name=sheet_names[1] if len(sheet_names) > 1 else sheet_names[0])

    st.sidebar.header("⚙️ Configurações")
    st.sidebar.markdown(f"**Planilha de Compras:** {sheet_names[0]}")
    st.sidebar.markdown(f"**Planilha de Negociação:** {sheet_names[1] if len(sheet_names) > 1 else sheet_names[0]}")

    codunds = compras["CODUND"].dropna().unique()
    atualizados = False
    for codun in codunds:
        codun_str = str(codun).upper()
        if codun_str not in FATOR_CONVERSAO:
            FATOR_CONVERSAO[codun_str] = 1
            atualizados = True
    if atualizados:
        salvar_fatores(FATOR_CONVERSAO)

    with st.sidebar.expander("✏️ Editar Fator de Conversão", expanded=False):
        with st.form("editar_fatores"):
            novos_fatores = {}
            for codun in sorted(FATOR_CONVERSAO.keys()):
                valor_atual = FATOR_CONVERSAO[codun]
                try:
                    valor_float = float(valor_atual if valor_atual not in [None, ''] else 1)
                except (ValueError, TypeError):
                    valor_float = 1.0
                    
                valor = st.number_input(f"{codun}", value=valor_float, step=0.01, format="%g")
                novos_fatores[codun] = valor
            
            salvar = st.form_submit_button("Salvar Fatores")
            
            if salvar:
                FATOR_CONVERSAO.update(novos_fatores)
                salvar_fatores(FATOR_CONVERSAO)
                st.success("Fatores de conversão salvos!")
                st.rerun()

    # Processamento
    dados = processar_dados(compras, negociacao)

    # Filtros
    colunas_disp = list(dados.columns)
    col_conta = next((c for c in colunas_disp if "CONTA" in c.upper() and ("CONTABIL" in c.upper() or "CONTÁBIL" in c.upper())), None)
    col_natureza = next((c for c in colunas_disp if "NATUR" in c.upper()), None)

    dados_filtrados = dados.copy()

    if "CODCOLIGADA" in dados_filtrados.columns:
        coligada_opcoes = sorted(dados_filtrados["CODCOLIGADA"].dropna().unique().astype(int).tolist())
        default_coligadas = [c for c in coligada_opcoes if c in [1, 11]]
        if not default_coligadas: default_coligadas = coligada_opcoes
        coligada_sel = st.sidebar.multiselect("Filtrar por Coligada:", coligada_opcoes, default=default_coligadas)
        dados_filtrados = dados_filtrados[dados_filtrados["CODCOLIGADA"].isin(coligada_sel)]
    else:
        st.sidebar.warning("Coluna 'CODCOLIGADA' não encontrada para filtro.")

    if col_natureza:
        natureza_opcoes = sorted(dados_filtrados[col_natureza].dropna().unique().tolist())
        natureza_sel = st.sidebar.multiselect("Filtrar por Natureza:", natureza_opcoes, default=natureza_opcoes)
        dados_filtrados = dados_filtrados[dados_filtrados[col_natureza].isin(natureza_sel)]

    if col_conta:
        conta_opcoes = sorted(dados_filtrados[col_conta].dropna().astype(str).unique().tolist())
        conta_sel = st.sidebar.multiselect("Filtrar por Conta Contábil:", conta_opcoes, default=conta_opcoes)
        dados_filtrados = dados_filtrados[dados_filtrados[col_conta].astype(str).isin(conta_sel)]

    produtos_opcoes = sorted(dados_filtrados["NOME_PRODUTO"].dropna().unique().tolist())
    produto_sel = st.sidebar.multiselect("Filtrar por Produto:", produtos_opcoes, default=produtos_opcoes)
    dados_filtrados = dados_filtrados[dados_filtrados["NOME_PRODUTO"].isin(produto_sel)]

    # Dashboards e Tabelas
    resumo = gerar_resumo(dados_filtrados)
    total_saving = dados_filtrados.loc[dados_filtrados["SAVING_POSITIVO"], "SAVING"].sum()
    total_produtos_saving = dados_filtrados.loc[dados_filtrados["SAVING_POSITIVO"], "CODIGO_PRD"].nunique()
    total_registros = len(dados_filtrados)

    col1, col2, col3 = st.columns(3)
    col1.metric("💰 Economias Alcançadas", f"R$ {total_saving:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    col2.metric("🏷️ Itens Negociados", total_produtos_saving)
    col3.metric("🧾 Registros Filtrados", total_registros)

    st.subheader("📋 Tabela Detalhada (Economias Apuradas)")
    mostrar_somente_saving = st.checkbox("Mostrar apenas saving > 0", value=True)
    resumo_exibir = resumo[resumo["SAVING"] > 0].copy() if mostrar_somente_saving else resumo.copy()
    resumo_fmt = resumo_exibir.copy()
    resumo_fmt["SAVING"] = resumo_fmt["SAVING"].apply(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    st.dataframe(resumo_fmt, use_container_width=True)

    saving_mensal = (
        dados_filtrados[dados_filtrados["SAVING_POSITIVO"]]
        .groupby("MÊS_ANO", as_index=False)["SAVING"].sum()
        .sort_values("MÊS_ANO")
    )

    fig = px.bar(
        saving_mensal, x="MÊS_ANO", y="SAVING", text_auto=".2s",
        title="💹 Evolução Mensal das Economias Alcançadas",
        labels={"SAVING": "Saving Total (R$)", "MÊS_ANO": "Mês/Ano"},
        color="SAVING", color_continuous_scale="Greens"
    )
    fig.update_layout(xaxis_tickangle=-45, height=500)
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("📈 Saving por Produto (Top 10)")
    saving_por_produto = (
        dados_filtrados[dados_filtrados["SAVING_POSITIVO"]]
        .groupby("NOME_PRODUTO", as_index=False)["SAVING"].sum()
        .sort_values("SAVING", ascending=False)
        .head(10)
    )

    if not saving_por_produto.empty:
        fig_prod = px.bar(
            saving_por_produto, x="NOME_PRODUTO", y="SAVING", text_auto=".2s",
            title="📊 Top 10 Produtos com Maiores Economias Alcançadas",
            labels={"SAVING": "Saving Total (R$)", "NOME_PRODUTO": "Produto"},
            color="SAVING", color_continuous_scale="Viridis"
        )
        fig_prod.update_layout(xaxis_tickangle=-45, height=500)
        st.plotly_chart(fig_prod, use_container_width=True)
    else:
        st.info("Nenhum saving positivo encontrado para exibir por produto.")

    st.subheader("🔮 Projeção de Saving por Produto (12 Meses)")
    st.markdown("""
    Esta projeção considera o consumo médio mensal e o saving médio por unidade, 
    baseado nos dados filtrados, para estimar o saving potencial nos próximos 12 meses.
    """)

    num_meses = dados_filtrados["MÊS_ANO"].nunique()
    if num_meses > 0:
        consumo_mensal_por_produto = (
            dados_filtrados.groupby("NOME_PRODUTO")["QTDENTRADA_UN"].sum() / num_meses
        ).reset_index()
        consumo_mensal_por_produto.rename(columns={"QTDENTRADA_UN": "CONSUMO_MEDIO_MENSAL"}, inplace=True)

        saving_por_unidade = (
            dados_filtrados[dados_filtrados["SAVING_POSITIVO"]]
            .groupby("NOME_PRODUTO")
            .apply(lambda x: x["SAVING"].sum() / x["QTDENTRADA_UN"].sum() if x["QTDENTRADA_UN"].sum() > 0 else 0)
            .reset_index(name="SAVING_MEDIO_POR_UNIDADE")
        )

        projecao_saving = pd.merge(consumo_mensal_por_produto, saving_por_unidade, on="NOME_PRODUTO", how="left").fillna(0)
        projecao_saving["PROJECAO_12_MESES"] = (
            projecao_saving["CONSUMO_MEDIO_MENSAL"] * projecao_saving["SAVING_MEDIO_POR_UNIDADE"] * 12
        )
        projecao_saving = projecao_saving[projecao_saving["PROJECAO_12_MESES"] > 0].sort_values("PROJECAO_12_MESES", ascending=False)

        projecao_fmt = projecao_saving.copy()
        for col in ["CONSUMO_MEDIO_MENSAL", "SAVING_MEDIO_POR_UNIDADE", "PROJECAO_12_MESES"]:
            projecao_fmt[col] = projecao_fmt[col].apply(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if "SAVING" in col or "PROJECAO" in col else f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        st.dataframe(projecao_fmt, use_container_width=True)
    else:
        st.info("Não há dados suficientes para gerar a projeção de saving.")
        projecao_saving = pd.DataFrame()

    # Exportação
    st.markdown("### 💾 Exportar Relatório Analítico")
    dados_exportar_resumo = resumo_exibir.copy()
    dados_exportar_projecao = projecao_saving.copy()
    
    col_excel, col_pdf = st.columns(2)

    excel_data = to_excel_bytes(dados_exportar_resumo)
    col_excel.download_button("⬇️ Baixar Resumo (Excel)", data=excel_data, file_name="resumo_saving.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    try:
        # Passamos também dados_filtrados para gerar o gráfico
        pdf_data = to_pdf_bytes(dados_exportar_resumo, dados_exportar_projecao, total_saving, total_produtos_saving, total_registros, dados_filtrados)
        col_pdf.download_button("📄 Baixar PDF", data=pdf_data, file_name="relatorio_saving.pdf", mime="application/pdf", use_container_width=True)
    except Exception as e_pdf:
        col_pdf.error(f"Erro ao gerar PDF: {e_pdf}")

except Exception as e:
    st.error(f"❌ Ocorreu um erro no processamento geral da aplicação:\n`{e}`")

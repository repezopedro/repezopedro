import pandas as pd
from datetime import timedelta
import numpy as np
import os
import json
from pandas.tseries.offsets import BDay
from functools import lru_cache

# ==============================================================================
# 1. Configurações e Funções Auxiliares
# ==============================================================================

# Definição do Diretório Base (Ajuste conforme o seu ambiente)
BASE_DIR = r'C:\Users\pedro.muniz\Documents\pythonlocal\BIOXXI\INDICADORESCMEXXFAB'

def add_business_days(date, days_to_add):
    """Calcula a data 'X' dias úteis após a data inicial (excluindo Sáb/Dom)."""
    if pd.isna(date):
        return pd.NaT
    
    current_date = date
    while days_to_add > 0:
        current_date += timedelta(days=1)
        if current_date.weekday() < 5:  # Se não for Sábado (5) ou Domingo (6)
            days_to_add -= 1  
    return current_date

@lru_cache(maxsize=1)
def load_data(file_path):
    """
    Carrega, filtra e prepara os dados da planilha Excel na aba correta.
    CORREÇÃO: Garante que CODCOLIGADA e CENTRO_CUSTO são strings.
    """
    SHEET_NAME = 'VIEW_PEDIDOS_SLA'
    try:
        df = pd.read_excel(file_path, sheet_name=SHEET_NAME)
    except FileNotFoundError:
        print(f"❌ ERRO: Arquivo não encontrado em: {file_path}. Por favor, verifique o caminho.")
        return None
    except ValueError:
        print(f"❌ ERRO: Planilha '{SHEET_NAME}' não encontrada no arquivo Excel.")
        return None
    except Exception as e:
        print(f"❌ ERRO ao carregar os dados: {e}")
        return None

    df.columns = df.columns.str.strip().str.upper()

    COLUMN_CMEXX = 'DATA_APROVACAO_CMEXX'
    if COLUMN_CMEXX not in df.columns:
        print(f"❌ ERRO CRÍTICO: Coluna '{COLUMN_CMEXX}' não encontrada.")
        return None

    # Colunas de Data
    date_cols = ['DTCRIACAO', 'DATACOMPETENCIA', 'DTAPROVACAO', 'DTAPROVACAOGERENCIA',
                 'DTIMPORTACAO', 'DATAEMISSAO', 'DTCONFIRMACAO', 'DTCONFIRMACAOSUPRIMENTOS',
                 COLUMN_CMEXX, 'DATA_PREVISTA_ENTREGA']
    # Colunas de ID que devem ser STRING para filtragem no JS
    id_cols = ['IDSOLICITACAO', 'NUMEROMOV', 'IDMOV', 'CENTRO_CUSTO', 'PEDIDO_EXTRA', 
               'STATUS_MOVIMENTO', 'CODCOLIGADA'] 
    
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
            
    # Tratamento da DATACOMPETENCIA
    if 'DATACOMPETENCIA' in df.columns and not df['DATACOMPETENCIA'].empty:
        df['DATACOMPETENCIA_BASE'] = df['DATACOMPETENCIA'].dt.to_period('M').dt.start_time
        df['DATACOMPETENCIA'] = df['DATACOMPETENCIA_BASE'] + BDay(0)
        df = df.drop(columns=['DATACOMPETENCIA_BASE'])
        
    # Converte IDs para String
    for col in id_cols:
          if col in df.columns:
            df[col] = df[col].astype(str)

    # Cria colunas de cálculo
    df = df[df['PEDIDO_EXTRA'] != 'S'].copy()
    df['AnoCriacao'] = df['DTCRIACAO'].dt.year
    df['MesAnoCompetencia'] = df['DATACOMPETENCIA'].dt.to_period('M')
    df['Data de Criação'] = df['DTCRIACAO'].dt.to_period('M').astype(str)
    
    df['DiaCriacao'] = df['DTCRIACAO'].dt.day
    df['DentroPrazoCorte'] = (df['DiaCriacao'] <= 20)

    # Cálculo SLAs
    df['SLA_CMEXX_Lim'] = df.apply(lambda row: add_business_days(row['DTCRIACAO'], 2), axis=1)
    df['SLA_CMEXX_Cumprido'] = df[COLUMN_CMEXX] <= df['SLA_CMEXX_Lim']
    df['SLA_CMEXX_Cumprido'] = df['SLA_CMEXX_Cumprido'].fillna(False)

    df['SLA_IMP_Lim'] = df.apply(lambda row: add_business_days(row[COLUMN_CMEXX], 3), axis=1)
    df['SLA_IMP_Cumprido'] = df['DTIMPORTACAO'] <= df['SLA_IMP_Lim']
    df['SLA_IMP_Cumprido'] = df['SLA_IMP_Cumprido'].fillna(False)

    df['Entregue_No_Prazo_Competencia'] = (df['DTCONFIRMACAOSUPRIMENTOS'] <= df['DATACOMPETENCIA']).fillna(False)

    return df

# ==============================================================================
# 2. Funções de Cálculo KPI
# ==============================================================================

def calculate_kpi_adesao_corte(df):
    if df.empty: return 0.0
    grouped = df.groupby('IDSOLICITACAO').first().reset_index()
    total_pedidos = grouped['IDSOLICITACAO'].nunique()
    pedidos_no_prazo = grouped[grouped['DentroPrazoCorte'] == True]['IDSOLICITACAO'].nunique()
    return (pedidos_no_prazo / total_pedidos) if total_pedidos > 0 else 0.0

def calculate_kpi_sla_cmexx(df):
    if df.empty: return 0.0
    grouped = df.groupby('IDSOLICITACAO').first().reset_index()
    total_pedidos = grouped['IDSOLICITACAO'].nunique()
    pedidos_no_prazo = grouped[grouped['SLA_CMEXX_Cumprido'] == True]['IDSOLICITACAO'].nunique()
    return (pedidos_no_prazo / total_pedidos) if total_pedidos > 0 else 0.0

def calculate_kpi_sla_importacao(df):
    if df.empty: return 0.0
    grouped = df.groupby('IDSOLICITACAO').first().reset_index()
    total_pedidos = grouped['IDSOLICITACAO'].nunique()
    pedidos_no_prazo = grouped[grouped['SLA_IMP_Cumprido'] == True]['IDSOLICITACAO'].nunique()
    return (pedidos_no_prazo / total_pedidos) if total_pedidos > 0 else 0.0

def calculate_kpi_entrega_competencia(df):
    if df.empty: return 0.0
    grouped = df.groupby('IDSOLICITACAO').first().reset_index()
    total_pedidos = grouped['IDSOLICITACAO'].nunique()
    pedidos_no_prazo = grouped[grouped['Entregue_No_Prazo_Competencia'] == True]['IDSOLICITACAO'].nunique()
    return (pedidos_no_prazo / total_pedidos) if total_pedidos > 0 else 0.0

# Funções de Dados Mensais (usadas para os gráficos)
def calculate_entrega_competencia_monthly_data(df):
    if df.empty: return pd.DataFrame()
    df_unique = df.groupby('IDSOLICITACAO').first().reset_index()
    df_unique['MesAnoCompetencia_str'] = df_unique['DATACOMPETENCIA'].dt.to_period('M').astype(str)
    grouped_final = df_unique.groupby('MesAnoCompetencia_str').agg(
        Total_Pedidos=('IDSOLICITACAO', 'nunique'),
        Pedidos_No_Prazo=('Entregue_No_Prazo_Competencia', 'sum')
    ).reset_index().sort_values('MesAnoCompetencia_str')
    grouped_final['Pedidos_Fora_Prazo'] = grouped_final['Total_Pedidos'] - grouped_final['Pedidos_No_Prazo']
    grouped_final['Percentual_SLA'] = (grouped_final['Pedidos_No_Prazo'] / grouped_final['Total_Pedidos']) * 100
    return grouped_final.fillna(0)

def calculate_adesao_by_month(df):
    if df.empty: return pd.DataFrame()
    df_unique = df.groupby('IDSOLICITACAO').first().reset_index()
    df_unique['MesAnoCriacao_str'] = df_unique['DTCRIACAO'].dt.to_period('M').astype(str)
    grouped = df_unique.groupby('MesAnoCriacao_str').agg(
        Total_Pedidos=('IDSOLICITACAO', 'nunique'),
        Pedidos_No_Prazo=('DentroPrazoCorte', 'sum')
    ).reset_index().sort_values('MesAnoCriacao_str')
    grouped['Percentual_Adesao'] = (grouped['Pedidos_No_Prazo'] / grouped['Total_Pedidos']) * 100
    return grouped.fillna(0)

def calculate_sla_cmexx_by_month(df):
    if df.empty: return pd.DataFrame()
    df_unique = df.groupby('IDSOLICITACAO').first().reset_index()
    df_unique['MesAnoCriacao_str'] = df_unique['DTCRIACAO'].dt.to_period('M').astype(str)
    grouped = df_unique.groupby('MesAnoCriacao_str').agg(
        Total_Pedidos=('IDSOLICITACAO', 'nunique'),
        Aprovacoes_No_Prazo=('SLA_CMEXX_Cumprido', 'sum')
    ).reset_index().sort_values('MesAnoCriacao_str')
    grouped['Percentual_SLA_CMEXX'] = (grouped['Aprovacoes_No_Prazo'] / grouped['Total_Pedidos']) * 100
    return grouped.fillna(0)

def calculate_sla_importacao_by_month(df):
    if df.empty: return pd.DataFrame()
    df_unique = df.groupby('IDSOLICITACAO').first().reset_index()
    df_unique['MesAnoCriacao_str'] = df_unique['DTCRIACAO'].dt.to_period('M').astype(str)
    grouped = df_unique.groupby('MesAnoCriacao_str').agg(
        Total_Pedidos=('IDSOLICITACAO', 'nunique'),
        Importacoes_No_Prazo=('SLA_IMP_Cumprido', 'sum')
    ).reset_index().sort_values('MesAnoCriacao_str')
    grouped['Percentual_SLA_Importacao'] = (grouped['Importacoes_No_Prazo'] / grouped['Total_Pedidos']) * 100
    return grouped.fillna(0)


# ==============================================================================
# 3. Função Principal de Processamento e Exportação
# ==============================================================================

def main_process():
    
    excel_file_name = 'CMEXXFAB_SQL.xlsx'
    file_path = os.path.join(BASE_DIR, excel_file_name)
    
    df_base = load_data(file_path)

    if df_base is None or df_base.empty:
        print("Processamento de dados interrompido devido a erro ou dados vazios.")
        return

    df_filtered = df_base.copy()
    
    # 1. Calcular KPIs Globais
    kpi_entrega = calculate_kpi_entrega_competencia(df_filtered)
    kpi_adesao = calculate_kpi_adesao_corte(df_filtered)
    kpi_cmexx = calculate_kpi_sla_cmexx(df_filtered)
    kpi_importacao = calculate_kpi_sla_importacao(df_filtered)

    # 2. Calcular Tendências Mensais (Gráficos)
    df_entrega_mensal = calculate_entrega_competencia_monthly_data(df_filtered)
    df_adesao_mensal = calculate_adesao_by_month(df_filtered)
    df_cmexx_mensal = calculate_sla_cmexx_by_month(df_filtered)
    df_importacao_mensal = calculate_sla_importacao_by_month(df_filtered)
    
    # 3. Estrutura da Tabela de Detalhes
    cols_table = ['IDSOLICITACAO', 'NUMEROMOV', 'CENTRO_CUSTO', 'DTCRIACAO', 'DATACOMPETENCIA',
                  'DTCONFIRMACAOSUPRIMENTOS', 'DATA_PREVISTA_ENTREGA', 'STATUS_MOVIMENTO']
    
    # 4. Listas para Filtros da Sidebar
    coligadas_list = sorted(df_filtered['CODCOLIGADA'].unique().tolist())
    anos_list = sorted(df_filtered['AnoCriacao'].unique().tolist(), reverse=True)
    centrais_list = sorted(df_filtered['CENTRO_CUSTO'].unique().tolist())

    cols_to_display = [col for col in cols_table if col in df_filtered.columns]
    
    # Prepara a Tabela de Detalhes (500 últimas linhas)
    df_table = df_filtered.groupby('IDSOLICITACAO').first().reset_index()[cols_to_display].sort_values('DTCRIACAO', ascending=False).head(500)
    
    # Formata as datas para exibição na tabela HTML
    df_table['DTCRIACAO'] = df_table['DTCRIACAO'].dt.strftime('%d/%m/%Y').fillna('')
    df_table['DATACOMPETENCIA'] = df_table['DATACOMPETENCIA'].dt.strftime('%d/%m/%Y').fillna('')
    df_table['DTCONFIRMACAOSUPRIMENTOS'] = df_table['DTCONFIRMACAOSUPRIMENTOS'].dt.strftime('%d/%m/%Y').fillna('')
    df_table['DATA_PREVISTA_ENTREGA'] = df_table['DATA_PREVISTA_ENTREGA'].dt.strftime('%d/%m/%Y').fillna('')
    
    # Prepara Dados Brutos para Filtragem dos KPIs no Frontend (apenas colunas necessárias)
    raw_data_export = df_filtered[['IDSOLICITACAO', 'CODCOLIGADA', 'AnoCriacao', 'CENTRO_CUSTO', 'DentroPrazoCorte', 'SLA_CMEXX_Cumprido', 'SLA_IMP_Cumprido', 'Entregue_No_Prazo_Competencia', 'DTCRIACAO', 'DATACOMPETENCIA', 'DTCONFIRMACAOSUPRIMENTOS']].groupby('IDSOLICITACAO').first().reset_index().to_dict(orient='records')

    # 5. Monta o dicionário de resultados
    results = {
        'timestamp': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S'),
        'total_pedidos': df_filtered['IDSOLICITACAO'].nunique(),
        'filters': {
            'coligadas': coligadas_list,
            'anos': anos_list,
            'centrais': centrais_list,
            # Filtro de Competência removido daqui para ser usado no DataTables
        },
        'global_kpis': {
            'entrega': round(kpi_entrega, 4),
            'adesao': round(kpi_adesao, 4),
            'cmexx': round(kpi_cmexx, 4),
            'importacao': round(kpi_importacao, 4)
        },
        'monthly_data': {
            'entrega': df_entrega_mensal[['MesAnoCompetencia_str', 'Percentual_SLA', 'Total_Pedidos']].to_dict(orient='records'),
            'adesao': df_adesao_mensal[['MesAnoCriacao_str', 'Percentual_Adesao']].to_dict(orient='records'),
            'cmexx': df_cmexx_mensal[['MesAnoCriacao_str', 'Percentual_SLA_CMEXX']].to_dict(orient='records'),
            'importacao': df_importacao_mensal[['MesAnoCriacao_str', 'Percentual_SLA_Importacao']].to_dict(orient='records')
        },
        'detail_table': df_table.to_dict(orient='records'),
        'raw_data': raw_data_export # Exporta o Dict
    }
    
    # 6. Salva o arquivo JavaScript
    js_file_name = 'dashboard_data.js'
    js_path = os.path.join(BASE_DIR, js_file_name)
    
    try:
        os.makedirs(BASE_DIR, exist_ok=True)
        
        json_string = json.dumps(results, ensure_ascii=False, indent=4, default=str)
        js_content = f"const dashboardData = {json_string};"
        
        with open(js_path, 'w', encoding='utf-8') as f:
            f.write(js_content)
            
        print(f"✅ Dados processados e salvos em: '{js_path}'.")
    except Exception as e:
        print(f"❌ ERRO ao salvar o arquivo JS: {e}")


if __name__ == '__main__':
    main_process()
#!/usr/bin/env python
# coding: utf-8

# In[ ]:


!pip install XlsxWriter


# In[13]:


import pandas as pd
import os
import numpy as np
from datetime import datetime
import xlsxwriter


# In[14]:


import sys
print("Caminho do Python em uso:", sys.executable)

# In[2]:


# Lista de tuplas contendo DENOM_CIA e CNPJ
lista_empresas = [
    ("Cyrela", "73.178.600/0001-18", "CYRE3"),
    ("Melnick Desenvolvimento Imobiliário", "12.181.987/0001-77", "MELK3"),
    ("Even Construtora e Incorporadora", "43.470.988/0001-65", "EVEN3"),
    ("Vulcabras", "50.926.955/0001-42", "VULC3"),
    ("Ultrapar", "33.256.439/0001-39", "UGPA3"),
    ("Hospital Mater Dei", "16.676.520/0001-59", "MATD3"),
    ("Eucatex", "56.643.018/0001-66", "EUCA4"),
    ("Ambipar", "12.648.266/0001-24", "AMBP3")
]

# In[3]:


# Caminho para os arquivos e diretório base
#caminho_arquivo = '/content/drive/MyDrive/empresas_if_cvm/empresas_cnpj/lista_empresas.csv'
caminho_salvar = r"C:\Users\User\Desktop\valuation_cvm\valuation_cvm_clone\dados_empresas"
diretorio_base = r"C:\Users\User\Desktop\valuation_cvm\valuation_cvm_clone\exports\rel_indicadores_financeiros"
# Abrindo os arquivos CSV
free_cash_flow = pd.read_csv(os.path.join(diretorio_base, 'free_cash_flow.csv'), sep=';')
free_cash_flow_trim = pd.read_csv(os.path.join(diretorio_base, 'free_cash_flow_trim.csv'), sep=';')
receita_liquida = pd.read_csv(os.path.join(diretorio_base, 'receita_liquida.csv'), sep=';')
receita_liquida_trim = pd.read_csv(os.path.join(diretorio_base, 'receita_liquida_trim.csv'), sep=';')
lucro_liquido = pd.read_csv(os.path.join(diretorio_base, 'lucro_liquido.csv'), sep=';')
lucro_liquido_trim = pd.read_csv(os.path.join(diretorio_base, 'lucro_liquido_trim.csv'), sep=';')
divida_liquida = pd.read_csv(os.path.join(diretorio_base, 'divida_liquida.csv'), sep=';')
divida_liquida_trim = pd.read_csv(os.path.join(diretorio_base, 'divida_liquida_trim.csv'), sep=';')
EBITDA = pd.read_csv(os.path.join(diretorio_base, 'EBITDA.csv'), sep=';')
EBITDA_trim = pd.read_csv(os.path.join(diretorio_base, 'EBITDA_trim.csv'), sep=';')
patrimonio_liquido = pd.read_csv(os.path.join(diretorio_base, 'patrimonio_liquido.csv'), sep=';')
patrimonio_liquido_trim = pd.read_csv(os.path.join(diretorio_base, 'patrimonio_liquido_trim.csv'), sep=';')
ativo_total = pd.read_csv(os.path.join(diretorio_base, 'ativo_total.csv'), sep=';')
ativo_total_trim = pd.read_csv(os.path.join(diretorio_base, 'ativo_total_trim.csv'), sep=';')
cx_liq_ativ_financ = pd.read_csv(os.path.join(diretorio_base, 'cx_liq_ativ_financ.csv'), sep=';')
cx_liq_ativ_financ_trim = pd.read_csv(os.path.join(diretorio_base, 'cx_liq_ativ_financ_trim.csv'), sep=';')
fcf_rl = pd.read_csv(os.path.join(diretorio_base, 'fcf_rl.csv'), sep=';')
fcf_rl_trim = pd.read_csv(os.path.join(diretorio_base, 'fcf_rl_trim.csv'), sep=';')
ll_sobre_rl = pd.read_csv(os.path.join(diretorio_base, 'll_sobre_rl.csv'), sep=';')
ll_sobre_rl_trim = pd.read_csv(os.path.join(diretorio_base, 'll_sobre_rl_trim.csv'), sep=';')
divida_sobre_ebitda = pd.read_csv(os.path.join(diretorio_base, 'divida_sobre_ebitda.csv'), sep=';')
divida_sobre_ebitda_trim = pd.read_csv(os.path.join(diretorio_base, 'divida_sobre_ebitda_trim.csv'), sep=';')
roe = pd.read_csv(os.path.join(diretorio_base, 'roe.csv'), sep=';')
roe_trim = pd.read_csv(os.path.join(diretorio_base, 'roe_trim.csv'), sep=';')
roa = pd.read_csv(os.path.join(diretorio_base, 'roa.csv'), sep=';')
roa_trim = pd.read_csv(os.path.join(diretorio_base, 'roa_trim.csv'), sep=';')
debt_to_equity = pd.read_csv(os.path.join(diretorio_base, 'debt_to_equity.csv'), sep=';')
debt_to_equity_trim = pd.read_csv(os.path.join(diretorio_base, 'debt_to_equity_trim.csv'), sep=';')
debt_to_total_assets = pd.read_csv(os.path.join(diretorio_base, 'debt_to_total_assets.csv'), sep=';')
debt_to_total_assets_trim = pd.read_csv(os.path.join(diretorio_base, 'debt_to_total_assets_trim.csv'), sep=';')
roic = pd.read_csv(os.path.join(diretorio_base, 'roic.csv'), sep=';')
roic_trim = pd.read_csv(os.path.join(diretorio_base, 'roic_trim.csv'), sep=';')
alavancagem_financeira = pd.read_csv(os.path.join(diretorio_base, 'alavancagem_financeira.csv'), sep=';')
alavancagem_financeira_trim = pd.read_csv(os.path.join(diretorio_base, 'alavancagem_financeira_trim.csv'), sep=';')
df_numero_acoes = pd.read_csv(os.path.join(diretorio_base, 'df_numero_acoes.csv'), sep=';')
fluxo_cx_investimentos = pd.read_csv(os.path.join(diretorio_base, 'fluxo_cx_investimentos.csv'), sep=';')
fluxo_cx_investimentos_trim=pd.read_csv(os.path.join(diretorio_base, 'fluxo_cx_investimentos_trim.csv'), sep=';')
fluxo_cx_operacional=pd.read_csv(os.path.join(diretorio_base, 'fluxo_cx_operacional.csv'), sep=';')
fluxo_cx_operacional_trim=pd.read_csv(os.path.join(diretorio_base, 'fluxo_cx_operacional_trim.csv'), sep=';')
margem_ebitda=pd.read_csv(os.path.join(diretorio_base, 'margem_ebitda.csv'), sep=';')
margem_ebitda_trim=pd.read_csv(os.path.join(diretorio_base, 'margem_ebitda_trim.csv'), sep=';')

# In[8]:


# Função para consolidar dados com base no CNPJ_CIA
def consolidar_dados(df_historico, df_trimestral, cnpj_empresa):
    df_historico_filtered = df_historico[df_historico['CNPJ_CIA'] == cnpj_empresa]
    df_trimestral_filtered = df_trimestral[df_trimestral['CNPJ_CIA'] == cnpj_empresa]
    df_consolidado = pd.merge(df_historico_filtered, df_trimestral_filtered, on='CNPJ_CIA', how='left', suffixes=('_hist', '_trim'))
    colunas_remover = ['index_hist', 'index_trim', 'DENOM_CIA_trim', 'index']
    df_consolidado = df_consolidado.drop(columns=colunas_remover, errors='ignore')
    return df_consolidado

# Função para adicionar a coluna Indicador na posição 0
def adicionar_indicador(df, indicador_nome):
    if not df.empty:
        df.insert(0, 'Indicador', indicador_nome)
    return df
# Função para obter a data inicial padrão
def obter_data_inicial(ticker):
    ano_atual = datetime.now().year
    data_inicial_padrao = datetime(ano_atual - 12, 1, 1).strftime('%Y-%m-%d')
    return data_inicial_padrao

# Função para baixar cotações
def baixar_cotacoes(ticker, nome_empresa):
    try:
        data_inicial_padrao = obter_data_inicial(ticker)
        cotacoes_empresa = yf.download(ticker, period="max", progress=False)
        if cotacoes_empresa.empty:
            return pd.DataFrame()
        primeira_data = cotacoes_empresa.index.min()
        ano_limite = datetime.now().year - 12
        data_limite = datetime(ano_limite, 1, 1)
        if primeira_data > data_limite:
            data_inicial = primeira_data.strftime('%Y-%m-%d')
        else:
            data_inicial = data_inicial_padrao
        cotacoes_empresa = yf.download(ticker, start=data_inicial, end=f'{datetime.now().year}-12-31', progress=False)
        if cotacoes_empresa.empty:
            return pd.DataFrame()
        cotacoes_empresa.reset_index(inplace=True)
        cotacoes_empresa['Nome Empresa'] = nome_empresa
        cotacoes_empresa['Ticker'] = ticker
        return cotacoes_empresa
    except Exception as e:
        print(f"Erro ao baixar cotações para {ticker}: {e}")
        return pd.DataFrame()

# Função para gerar a tabela DY
def gerar_tabela_dy(ativo_total):
    colunas_remover = ['index_hist', 'DENOM_CIA_hist', 'CNPJ_CIA', 'index_trim', 'DENOM_CIA_trim']
    ativo_total_clean = ativo_total.drop(columns=colunas_remover, errors='ignore')
    colunas_dy = ativo_total_clean.columns
    dy_data = {coluna: [None] for coluna in colunas_dy}
    df_dy = pd.DataFrame(dy_data)
    #df_dy.insert(0, 'Indicador', 'DY(%)')
    return df_dy

# Função principal
def consolidar_para_empresa(empresa_nome, empresa_cnpj, empresa_ticker, caminho_salvar):
    # Consolidar os dados de cada indicador
    ativo_total_consolidado = adicionar_indicador(consolidar_dados(ativo_total, ativo_total_trim, empresa_cnpj), 'Ativo Total')
    receita_liquida_consolidado = adicionar_indicador(consolidar_dados(receita_liquida, receita_liquida_trim, empresa_cnpj), 'Receita Líquida')
    lucro_liquido_consolidado = adicionar_indicador(consolidar_dados(lucro_liquido, lucro_liquido_trim, empresa_cnpj), 'Lucro Líquido')
    divida_liquida_consolidado = adicionar_indicador(consolidar_dados(divida_liquida, divida_liquida_trim, empresa_cnpj), 'Dívida Líquida')
    EBITDA_consolidado = adicionar_indicador(consolidar_dados(EBITDA, EBITDA_trim, empresa_cnpj), 'EBITDA')
    patrimonio_liquido_consolidado = adicionar_indicador(consolidar_dados(patrimonio_liquido, patrimonio_liquido_trim, empresa_cnpj), 'Patrimônio Líquido')
    cx_liq_ativ_financ_consolidado = adicionar_indicador(consolidar_dados(cx_liq_ativ_financ, cx_liq_ativ_financ_trim, empresa_cnpj), 'Caixa Líquido Atividades Financeiras')
    fluxo_cx_investimentos_consolidado = adicionar_indicador(consolidar_dados(fluxo_cx_investimentos, fluxo_cx_investimentos_trim, empresa_cnpj), 'Fluxo Caixa Investimentos')
    fluxo_cx_operacional_consolidado = adicionar_indicador(consolidar_dados(fluxo_cx_operacional, fluxo_cx_operacional_trim, empresa_cnpj), 'Fluxo Caixa Operacional')
    margem_ebitda_consolidado = adicionar_indicador(consolidar_dados(margem_ebitda, margem_ebitda_trim, empresa_cnpj), 'Margem EBITDA')
    # Certifique-se de que as colunas estão alinhadas e converta para numérico
    colunas_numericas = fluxo_cx_operacional_consolidado.columns[2:]  # Supondo que as primeiras duas colunas são 'CNPJ_CIA' e 'Indicador'

    # Criar uma cópia para evitar alterar os DataFrames originais
    fcf_consolidado = fluxo_cx_operacional_consolidado.copy()
    for coluna in colunas_numericas:
      fluxo_operacional_valores = pd.to_numeric(fluxo_cx_operacional_consolidado[coluna], errors='coerce')
      caixa_ativ_financ_valores = pd.to_numeric(cx_liq_ativ_financ_consolidado[coluna], errors='coerce')
      # Somar as colunas correspondentes
      fcf_consolidado[coluna] = fluxo_operacional_valores + caixa_ativ_financ_valores
      # Atualizar o nome do indicador
      fcf_consolidado['Indicador'] = 'FCF'

    # Multiplicar as colunas numéricas de cada DataFrame por 1000
    ativo_total_consolidado.iloc[:, 2:] = ativo_total_consolidado.iloc[:, 2:].apply(pd.to_numeric, errors='coerce') * 1000
    receita_liquida_consolidado.iloc[:, 2:] = receita_liquida_consolidado.iloc[:, 2:].apply(pd.to_numeric, errors='coerce') * 1000
    lucro_liquido_consolidado.iloc[:, 2:] = lucro_liquido_consolidado.iloc[:, 2:].apply(pd.to_numeric, errors='coerce') * 1000
    divida_liquida_consolidado.iloc[:, 2:] = divida_liquida_consolidado.iloc[:, 2:].apply(pd.to_numeric, errors='coerce') * 1000
    EBITDA_consolidado.iloc[:, 2:] = EBITDA_consolidado.iloc[:, 2:].apply(pd.to_numeric, errors='coerce') * 1000
    patrimonio_liquido_consolidado.iloc[:, 2:] = patrimonio_liquido_consolidado.iloc[:, 2:].apply(pd.to_numeric, errors='coerce') * 1000
    cx_liq_ativ_financ_consolidado.iloc[:, 2:] = cx_liq_ativ_financ_consolidado.iloc[:, 2:].apply(pd.to_numeric, errors='coerce') * 1000
    fluxo_cx_investimentos_consolidado.iloc[:, 2:] = fluxo_cx_investimentos_consolidado.iloc[:, 2:].apply(pd.to_numeric, errors='coerce') * 1000
    fluxo_cx_operacional_consolidado.iloc[:, 2:] = fluxo_cx_operacional_consolidado.iloc[:, 2:].apply(pd.to_numeric, errors='coerce') * 1000
    fcf_consolidado.iloc[:, 2:] = fcf_consolidado.iloc[:, 2:].apply(pd.to_numeric, errors='coerce') * 1000

    # Consolidar outros indicadores
    ll_sobre_rl_consolidado = adicionar_indicador(consolidar_dados(ll_sobre_rl, ll_sobre_rl_trim, empresa_cnpj), 'Lucro Líquido sobre Receita Líquida')
    divida_sobre_ebitda_consolidado = adicionar_indicador(consolidar_dados(divida_sobre_ebitda, divida_sobre_ebitda_trim, empresa_cnpj), 'Dívida sobre EBITDA')
    roe_consolidado = adicionar_indicador(consolidar_dados(roe, roe_trim, empresa_cnpj), 'ROE')
    roa_consolidado = adicionar_indicador(consolidar_dados(roa, roa_trim, empresa_cnpj), 'ROA')
    debt_to_equity_consolidado = adicionar_indicador(consolidar_dados(debt_to_equity, debt_to_equity_trim, empresa_cnpj), 'Debt to Equity')
    debt_to_total_assets_consolidado = adicionar_indicador(consolidar_dados(debt_to_total_assets, debt_to_total_assets_trim, empresa_cnpj), 'Debt to Total Assets')
    roic_consolidado = adicionar_indicador(consolidar_dados(roic, roic_trim, empresa_cnpj), 'ROIC')
    alavancagem_financeira_consolidado = adicionar_indicador(consolidar_dados(alavancagem_financeira, alavancagem_financeira_trim, empresa_cnpj), 'Alavancagem Financeira')

    # Criar o DataFrame consolidado com todos os indicadores
    indicadores_list = [
        ativo_total_consolidado,
        receita_liquida_consolidado,
        lucro_liquido_consolidado,
        divida_liquida_consolidado,
        EBITDA_consolidado,
        patrimonio_liquido_consolidado,
        cx_liq_ativ_financ_consolidado,
        fluxo_cx_investimentos_consolidado,
        fluxo_cx_operacional_consolidado,
        fcf_consolidado,
        margem_ebitda_consolidado,
        ll_sobre_rl_consolidado,
        divida_sobre_ebitda_consolidado,
        roe_consolidado,
        roa_consolidado,
        debt_to_equity_consolidado,
        debt_to_total_assets_consolidado,
        roic_consolidado,
        alavancagem_financeira_consolidado
    ]

    # Verificar se todos os DataFrames são não vazios
    for df in indicadores_list:
        if df.empty:
            print("Um dos DataFrames está vazio. Verifique a consolidação dos dados.")
            return

    df_indicadores = pd.concat(indicadores_list, ignore_index=True)

    # Gerar a tabela DY com base na estrutura de colunas do ativo_total
    df_dy = gerar_tabela_dy(ativo_total_consolidado)

    # Baixar cotações
    cotacoes_empresa = baixar_cotacoes(empresa_ticker, empresa_nome)

    # Filtrar df_numero_acoes para a empresa específica
    df_numero_acoes_empresa = df_numero_acoes[df_numero_acoes['CNPJ_Companhia'] == empresa_cnpj]

    # Lista de colunas a serem removidas, se existirem
    colunas_remover = ['level_0_hist', 'level_0_trim', 'level_0']

    # Remover as colunas dos DataFrames que serão salvos
    df_indicadores = df_indicadores.drop(columns=colunas_remover, errors='ignore')
    df_dy = df_dy.drop(columns=colunas_remover, errors='ignore')
    cotacoes_empresa = cotacoes_empresa.drop(columns=colunas_remover, errors='ignore')
    df_numero_acoes_empresa = df_numero_acoes_empresa.drop(columns=colunas_remover, errors='ignore')

    # Salvar no Excel
    try:
      caminho_arquivo_excel = os.path.join(caminho_salvar, f'{empresa_nome}_{empresa_ticker}.xlsx')
      with pd.ExcelWriter(caminho_arquivo_excel, engine='xlsxwriter') as writer:
        df_indicadores.to_excel(writer, sheet_name='Indicadores', index=False)
        df_dy.to_excel(writer, sheet_name='DY(%)', index=False)
        if not cotacoes_empresa.empty:
            cotacoes_empresa.to_excel(writer, sheet_name='Cotacoes', index=False)
        if not df_numero_acoes_empresa.empty:
            df_numero_acoes_empresa.to_excel(writer, sheet_name='Numero de Acoes', index=False)
      print(f"Arquivo consolidado salvo com sucesso para {empresa_nome} ({empresa_ticker}) com cotação, DY(%) e número de ações!")
    except Exception as e:
      print(f"Erro ao salvar o arquivo para {empresa_nome} ({empresa_ticker}): {e}")

# Exemplo de uso
for empresa in lista_empresas:
    nome_empresa = empresa[0]
    cnpj_empresa = empresa[1]
    ticker_empresa = empresa[2] + ".SA"
    consolidar_para_empresa(nome_empresa, cnpj_empresa, ticker_empresa, caminho_salvar)


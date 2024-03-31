import pandas as pd

# Ajusta formatação de exibição do dataframe pandas
pd.options.display.float_format = '{:.2f}'.format

# Cria os dataframes para cada uma das abas da planilha Alura_PlanilhaDados.xlsx

## aba Principal
df_principal = pd.read_excel('Alura_PlanilhaDados.xlsx', 
                             sheet_name='Principal')
df_principal = df_principal.rename(columns={'Ativo': 'ativo',
                                            'Data': 'data',
                                            'Último (R$)': 'valor_final',
                                            'Var. Dia (%)': 'variacao_dia_pct'})

## aba Total_de_acoes
df_total_acoes = pd.read_excel('Alura_PlanilhaDados.xlsx', 
                               sheet_name='Total_de_acoes')
df_total_acoes = df_total_acoes.rename(columns={'Código': 'codigo_ticker', 
                                                'Qtde. Teórica': 'qtd_teorica'})

## aba Ticker
df_ticker = pd.read_excel('Alura_PlanilhaDados.xlsx', sheet_name='Ticker')
df_ticker = df_ticker.rename(columns={'Ticker': 'codigo_ticker', 
                                      'Nome': 'nome_ticker'})

## aba Segmento_Empresas_ChatGPT
df_segmento = pd.read_excel('Alura_PlanilhaDados.xlsx', 
                            sheet_name='Segmento_Empresas_ChatGPT')
df_segmento = df_segmento.rename(columns={'Empresa': 'nome_empresa', 
                                          'Segmento': 'segmento_empresa',
                                          'Idade (anos)': 'idade_empresa'})

################################################################################

# Manipulando dataframe com a aba Principal
df_principal = df_principal[['ativo', 'data', 'valor_final', 
                             'variacao_dia_pct']].copy()

# Adiciona coluna Variacao_dia_dec com o valor de variacao_dia_pct em 
# formato decimal
df_principal['variacao_dia_dec'] = df_principal['variacao_dia_pct'] / 100

# Calcula valor inicial do dia
df_principal['valor_inicial'] = df_principal['valor_final'] / \
                                (df_principal['variacao_dia_dec'] + 1)

# Faz o join com a aba total_acoes para recuperar qtd_teorica
df_principal = df_principal.merge(df_total_acoes, left_on='ativo', 
                                  right_on='codigo_ticker', how='left')
df_principal = df_principal.drop(columns=['codigo_ticker'])
df_principal['qtd_teorica'] = df_principal['qtd_teorica'].astype(int)

# Calcula variacao diaria em reais
df_principal['variacao_reais'] = (df_principal['valor_final'] - \
                                 df_principal['valor_inicial']) * \
                                 df_principal['qtd_teorica']

# Cria coluna com resultado_variacao: subiu, desceu, estável
df_principal['resultado_variacao'] = df_principal['variacao_reais'].apply(
    lambda x: 'Subiu' if x > 0 else ('Desceu' if x < 0 else 'Estavel'))

# Faz o join com a aba ticker para recuperar o nome da empresa da ação
df_principal = df_principal.merge(df_ticker, left_on='ativo', 
                                  right_on='codigo_ticker', how='left')
df_principal = df_principal.drop(columns=['codigo_ticker'])

# Faz o join com a aba ticker para recuperar o segmento e diade da empresa
df_principal = df_principal.merge(df_segmento, left_on='nome_ticker', 
                                  right_on='nome_empresa', how='left')
df_principal = df_principal.drop(columns=['nome_empresa'])

# Cria coluna com a classificao por idade da empresa
df_principal['categoria_idade'] = df_principal['idade_empresa'].apply(
    lambda x: 'Mais de 100 anos' if x > 100 
    else ('Menos de 50 anos' if x < 50 else 'Entre 50 e 100 anos'))

# Escreve dataframe resultante
print('################################################################################')
print('#                               DADOS ANALITICOS                               #')
print('################################################################################')
print('')
print(df_principal)
print('')

################################################################################

# Calcula apontamentos
print('################################################################################')
print('#                                 APONTAMENTOS                                 #')
print('################################################################################')
print('')
maior_variacao_reais = df_principal['variacao_reais'].max()
print(f"Valor maior variacao (R$):\t {maior_variacao_reais:.2f}")
menor_variacao_reais = df_principal['variacao_reais'].min()
print(f"Valor menor variacao (R$):\t {menor_variacao_reais:.2f}")
media_variacao_reais = df_principal['variacao_reais'].mean()
print(f"Valor média de variacao (R$):\t {media_variacao_reais:.2f}")

media_variacao_positiva = df_principal[df_principal['resultado_variacao'] == 'Subiu']['variacao_reais'].mean()
print(f"Valor média de variacao de quem subiu (R$):\t {media_variacao_positiva:.2f}")
media_variacao_negativa = df_principal[df_principal['resultado_variacao'] == 'Desceu']['variacao_reais'].mean()
print(f"Valor média de variacao de quem desceu (R$):\t {media_variacao_negativa:.2f}")
print('')

################################################################################

# Cria dataframe só com ações que subiram
df_subiu = df_principal[df_principal['resultado_variacao'] == 'Subiu']

# Cria dataframe agrupado por segemento das ações que subiram
df_analise_segmento_subiu = df_subiu.groupby('segmento_empresa')['variacao_reais'].sum().reset_index()

print('################################################################################')
print('#                         AGRUPAMENTO AÇÕES QUE SUBIRAM                        #')
print('################################################################################')
print('')
print(df_analise_segmento_subiu)
print('')
################################################################################

# Cria dataframe agrupado por segemento das ações que subiram
df_analise_saldo = df_principal.groupby('resultado_variacao')['variacao_reais'].sum().reset_index()

print('################################################################################')
print('#                  AGRUPAMENTO AÇÕES POR RESULTADO DA VARIAÇÃO                 #')
print('################################################################################')
print('')
print(df_analise_saldo)
print('')
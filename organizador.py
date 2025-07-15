import pandas as pd

# Carrega o Excel original e salva como CSV
df = pd.read_excel('./planilha/produtos25jun11.xlsx')
df.to_csv('./planilha/produtos25jun11.csv', index=False, encoding='utf-8')

# Recarrega o CSV (tratando tipos e limpeza)
df = pd.read_csv('./planilha/produtos25jun11.csv')

# Remove colunas desnecessárias
colunas_para_remover = [
    'ID', 'Unidade', 'Classificação fiscal', 'Origem', 'Preço', 'Valor IPI fixo', 'Observações', 'Situação', 'Estoque',
    'Preço de custo', 'Localização', 'Estoque máximo', 'Estoque mínimo', 'GTIN/EAN', 'GTIN/EAN tributável', 'CEST',
    'Código de Enquadramento IPI', 'Formato embalagem', 'Tipo do produto', 'URL imagem 1', 'URL imagem 2', 'URL imagem 3',
    'URL imagem 4', 'URL imagem 5', 'URL imagem 6', 'Categoria', 'Marca', 'Garantia', 'Sob encomenda', 'Preço promocional',
    'URL imagem externa 1', 'URL imagem externa 2', 'URL imagem externa 3', 'URL imagem externa 4', 'URL imagem externa 5',
    'URL imagem externa 6', 'Link do vídeo', 'Título SEO', 'Descrição SEO', 'Palavras chave SEO', 'Slug',
    'Dias para preparação', 'Controlar lotes', 'Unidade por caixa', 'URL imagem externa 7', 'URL imagem externa 8',
    'URL imagem externa 9', 'URL imagem externa 10', 'Markup', 'Permitir inclusão nas vendas', 'EX TIPI'
]
df.drop(columns=colunas_para_remover, inplace=True)

# Define a categoria dos produtos


def classificar_categoria(row):
    sku = str(row['Código (SKU)'])
    codigo_pai = str(row['Código do pai'])
    variacao = str(row['Variações'])

    if '.KIT' in sku:
        return 'KIT'
    elif '.V0' in sku:
        return 'VARIAÇÃO - PAI'
    elif codigo_pai != 'nan' and variacao != 'nan' and codigo_pai != '' and variacao != '':
        return 'VARIAÇÃO - FILHO'
    else:
        return 'SEM VARIAÇÃO'


df['Categoria'] = df.apply(classificar_categoria, axis=1)

# Define prioridade para ordenação
df['Prioridade'] = df['Categoria'].map({
    'SEM VARIAÇÃO': 1,
    'VARIAÇÃO - PAI': 2,
    'VARIAÇÃO - FILHO': 3,
    'KIT': 4
}).fillna(5)

# Define SKU Pai
df['SKU Pai'] = df.apply(lambda row: row['Código do pai'] if row['Categoria']
                         == 'VARIAÇÃO - FILHO' else row['Código (SKU)'], axis=1)

# Ordena
df_ordenado = df.sort_values(by=['Prioridade', 'SKU Pai', 'Variações'])

# Exporta
df_ordenado.to_excel('./planilha/planilha_organizada.xlsx', index=False)
print("Planilha organizada exportada com sucesso.")

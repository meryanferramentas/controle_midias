import pandas as pd

df = pd.read_excel('./planilha/planilha_organizada.xlsx')

renomear_fornecedores = {
    "BAZZI COMPANY COM IMP E EXP DE PRODUTOS ELETRÔNICOS EIRELI": "X-CELL",
    "RIO CHENS IMPORT. E EXPORTAD. LTDA": "BESTFER",
    "STARTOOLS FERRAMENTAS, COMERCIO, IMPORTACAO E EXPORTACAO LTD": "STARTOOLS",
    "TF TOP FUSION IND. DE TUBOS E CON. LTDA": "TOP FUSION",
    "TOP RIO COMERCIAL LTDA": "TOP PROS",
    "C3B COMERCIO DE IMPORTAÇÃO E EXPORTAÇÃO LTDA": "HUTZ"
}
df['Fornecedor'] = df['Fornecedor'].replace(renomear_fornecedores)

siglas_sku = {
    'BFH': 'BESTFER',
    'HTZ': 'HUTZ',
    'STR': 'STARTOOLS',
    'TFN': 'TOP FUSION',
    'TPR': 'TOP PROS',
    'XCL': 'X-CELL'
}


def corrigir_fornecedor(row):
    sku = str(row['Código (SKU)']).upper()
    if 'KIT' in sku:
        return row['Fornecedor']
    for sigla, fornecedor in siglas_sku.items():
        if sigla in sku:
            return fornecedor
    return row['Fornecedor']


df['Fornecedor'] = df.apply(corrigir_fornecedor, axis=1)


def is_valido(sku):
    sku = str(sku).upper()
    return 'MLB' not in sku and sku != '187939571889'


df = df[df['Código (SKU)'].apply(is_valido)]


def separar_descricao_variacao(descricao):
    if pd.isna(descricao):
        return '', ''
    partes = descricao.split(' - ', 1)
    if len(partes) == 2:
        return partes[0].strip(), partes[1].strip()
    return descricao.strip(), ''


descricoes = []
variacoes = []

for _, row in df.iterrows():
    if row['Categoria'] in ['VARIAÇÃO - PAI', 'VARIAÇÃO - FILHO', 'KIT']:
        desc, variacao = separar_descricao_variacao(row['Descrição'])
    else:
        desc, variacao = row['Descrição'], ''
    descricoes.append(desc)
    variacoes.append(variacao)

df['Descrição'] = descricoes
df['Descrição da Variação'] = variacoes

df['SKU Pai'] = df['Código do pai']
df['Código SKU'] = df['Código (SKU)']
df['Código do Fornecedor'] = df['Cód do Fornecedor']
df['Peso líquido'] = df['Peso líquido (Kg)']
df['Peso bruto'] = df['Peso bruto (Kg)']
df['Largura'] = df['Largura embalagem']
df['Altura'] = df['Altura embalagem']
df['Comprimento'] = df['Comprimento embalagem']

sem_variacao = df[df['Categoria'] == 'SEM VARIAÇÃO']
kits = df[df['Categoria'] == 'KIT']
pais = df[df['Categoria'] == 'VARIAÇÃO - PAI']
filhos = df[df['Categoria'] == 'VARIAÇÃO - FILHO']

resultado = []

todos_fornecedores = sorted(df['Fornecedor'].dropna().astype(str).unique())

for fornecedor in todos_fornecedores:
    fornecedor_df = df[df['Fornecedor'] == fornecedor]

    sem_var = fornecedor_df[fornecedor_df['Categoria'] == 'SEM VARIAÇÃO']
    resultado.append(sem_var)

    pais = fornecedor_df[fornecedor_df['Categoria'] == 'VARIAÇÃO - PAI']
    filhos = fornecedor_df[fornecedor_df['Categoria'] == 'VARIAÇÃO - FILHO']

    for _, pai in pais.iterrows():
        filhos_do_pai = filhos[filhos['Código do pai'] == pai['Código (SKU)']]
        if not filhos_do_pai.empty:
            resultado.append(filhos_do_pai)

    kits = fornecedor_df[fornecedor_df['Categoria'] == 'KIT']
    resultado.append(kits)

df_final = pd.concat(resultado, ignore_index=True)

df_final['Tem Peso'] = df_final[['Peso líquido', 'Peso bruto']
                                ].notna().any(axis=1).map({True: 'Sim', False: 'Não'})
df_final['Tem Dimensões'] = df_final[['Largura', 'Altura', 'Comprimento']
                                     ].notna().all(axis=1).map({True: 'Sim', False: 'Não'})
df_final['Tem Descrição Complementar'] = df_final['Descrição complementar'].notna(
).map({True: 'Sim', False: 'Não'})

colunas_finais = [
    'Fornecedor', 'Descrição', 'SKU Pai', 'Código SKU', 'Código do Fornecedor',
    'Descrição da Variação',
    'Tem Peso', 'Tem Dimensões', 'Tem Descrição Complementar'
]

df_final = df_final[colunas_finais]

df_final.to_excel('./planilha/planilha_formatada.xlsx', index=False)
print("Planilha final formatada salva com sucesso: planilha_formatada.xlsx")

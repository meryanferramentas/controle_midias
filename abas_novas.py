import pandas as pd

df = pd.read_excel('./planilha/planilha_formatada.xlsx')

colunas_para_preencher = ['Fornecedor', 'Descrição', 'SKU Pai']
df[colunas_para_preencher] = df[colunas_para_preencher].ffill()


def limpar_nome_aba(nome):
    caracteres_proibidos = ['\\', '/', '*', '?', ':', '[', ']']
    for c in caracteres_proibidos:
        nome = nome.replace(c, '')
    return nome[:31]


with pd.ExcelWriter('./planilha/planilha_com_abas.xlsx', engine='openpyxl') as writer:
    fornecedores_unicos = df['Fornecedor'].dropna().unique()

    for fornecedor in sorted(fornecedores_unicos):
        aba_nome = limpar_nome_aba(str(fornecedor))
        aba_df = df[df['Fornecedor'] == fornecedor]
        aba_df.to_excel(writer, sheet_name=aba_nome, index=False)

print("Arquivo com abas por fornecedor criado com sucesso: planilha_com_abas.xlsx")

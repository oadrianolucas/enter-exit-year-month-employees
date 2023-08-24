import pandas as pd
nome_arquivo = 'data.xlsx'  
planilha = pd.ExcelFile(nome_arquivo)

primeiro_ultimo_mes_ano_funcao = {}

for sheet_name in planilha.sheet_names:
    df = planilha.parse(sheet_name)
    if 'Nome' in df.columns:  
        for index, row in df.iterrows():
            nome = row['Nome']
            if pd.notna(nome):  
                if nome not in primeiro_ultimo_mes_ano_funcao:
                    primeiro_ultimo_mes_ano_funcao[nome] = {'Função': row['Função'],
                                                            'Primeiro_Ano': row['Ano'],
                                                            'Primeiro_Mês': row['Mês'],
                                                            'Último_Ano': row['Ano'],
                                                            'Último_Mês': row['Mês']}
                if pd.notna(row['Mês']):  
                    primeiro_ultimo_mes_ano_funcao[nome]['Último_Ano'] = row['Ano']
                    primeiro_ultimo_mes_ano_funcao[nome]['Último_Mês'] = row['Mês']

dados = []
for nome, info in primeiro_ultimo_mes_ano_funcao.items():
    dados.append({'Função': info['Função'],
                  'Nome': nome,
                  'Primeiro_Ano': info['Primeiro_Ano'],
                  'Primeiro_Mês': info['Primeiro_Mês'],
                  'Último_Ano': info['Último_Ano'],
                  'Último_Mês': info['Último_Mês']})
df_final = pd.DataFrame(dados)

nome_arquivo_xlsx = 'primeiro_ultimo_mes_ano_funcao.xlsx'
df_final.to_excel(nome_arquivo_xlsx, index=False)

print(f"Dados dos primeiros, últimos meses e anos e funções salvos em {nome_arquivo_xlsx}")

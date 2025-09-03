import pandas as pd

planilha_1 = pd.read_excel('publicacoes_legalone.xlsx')
planilha_2 = pd.read_excel('recorte_oab.xlsx')

planilha_2 = planilha_2.rename(columns={
    'Número do Processo': 'Número',
    'partes': 'Partes',
    'Advogado': 'Responsável principal',
})

planilha_final = pd.concat([planilha_1, planilha_2], ignore_index=True)

planilha_final = planilha_final.drop_duplicates(subset=['Número'], keep='first')

planilha_final.to_excel('publicacoes_final.xlsx', index=False)

print("Planilha final de Publicações gerada com sucesso.")

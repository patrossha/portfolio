import pandas as pd

# Carregar planilhas de andamentos
planilha_1 = pd.read_excel('andamentos_legalone.xlsx')  
planilha_2 = pd.read_excel('andamentos_pje.xlsx')  
planilha_3 = pd.read_excel('andamentos_eproc.xlsx')  
planilha_4 = pd.read_excel('andamentos_esaj.xlsx')
planilha_5 = pd.read_excel('andamentos_trt.xlsx')  
planilha_6 = pd.read_excel('andamentos_dcp.xlsx')  

# Renomear colunas
planilha_2 = planilha_2.rename(columns={
    'Número do Processo': 'Número',
    'Movimentos': 'Descrição',
})
planilha_3 = planilha_3.rename(columns={
    'Número do Processo': 'Número',
    'Movimentação': 'Descrição',
})
planilha_4 = planilha_4.rename(columns={
    'Número do Processo': 'Número',
    'Movimentos': 'Descrição',
})
planilha_5 = planilha_5.rename(columns={
    'Número do Processo': 'Número',
    'Eventos': 'Descrição',
})
planilha_6 = planilha_6.rename(columns={
    'Processo': 'Número',
    'Fase Atual': 'Descrição',
})

# concat
planilha_final = pd.concat([
    planilha_1, planilha_2, planilha_3,
    planilha_4, planilha_5, planilha_6
], ignore_index=True)

# drop duplicates
planilha_final = planilha_final.drop_duplicates(subset=['Número'], keep='first')

# Carregar a planilha Processos20251// atualizar 1xsem
processos_df = pd.read_excel('Processos20251.xlsx')

processos_df['cliente_nome'] = processos_df['Cliente principal'].apply(lambda x: x.split()[0] if isinstance(x, str) and x else '')
processos_df['contrario_nome'] = processos_df['Contrário principal'].apply(lambda x: x.split()[0] if isinstance(x, str) and x else '')
processos_df['Advogado'] = processos_df['Advogado'].apply(lambda x: x.split()[0] if isinstance(x, str) and x else '')
processos_df['partes'] = processos_df.apply(
    lambda row: f"{row['cliente_nome']} X {row['contrario_nome']}" if row['cliente_nome'] and row['contrario_nome'] else '',
    axis=1
)

processos_aux = processos_df[['Número do Processo', 'partes', 'Advogado']].copy()
processos_aux.rename(columns={
    'Número do Processo': 'Número',
    'partes': 'Partes',
    'Advogado': 'Responsável'
}, inplace=True)

# drop duplicates no DataFrame de processos_aux 
processos_aux = processos_aux.drop_duplicates(subset=['Número'], keep='first')

planilha_final['Partes'] = planilha_final['Número'].map(processos_aux.set_index('Número')['Partes'])
planilha_final['Responsável'] = planilha_final['Número'].map(processos_aux.set_index('Número')['Responsável'])

planilha_final.to_excel('andamentos_final.xlsx', index=False)

print("Planilha final de Andamentos gerada com sucesso")

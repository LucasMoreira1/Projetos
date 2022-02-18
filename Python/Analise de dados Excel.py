import pandas as pd


x = pd.read_excel('leads.xlsx',engine='openpyxl')

x.info()
x = x.drop(columns=['ID','Equipe','Data de chegada','Data de visualização'])

#indexar para realizar filtros
contatos = pd.read_excel('leads.xlsx',engine='openpyxl', index_col='Motivo de perda')

#Filtrar e exportar excel
filtro = contatos[contatos.index.isin(['Sem interação', 'Outros'])]

filtro = filtro.loc[:,['Nome do cliente', 'Telefone']]

filtro.to_excel('filtro.xlsx',index=False)

#converter resultado número em string para manipular os dados
automacao = pd.read_excel('filtro.xlsx',engine='openpyxl')

automacao.head(3)
nome = automacao['Nome do cliente']
telefone = automacao['Telefone'].astype(str)
telefone1 = telefone.str[-11:] #função (DIREITO do Excel) no exemplo mantendo ultimos 11 caracteres. 
automacao['Telefone'] = telefone1 #para substituir a planilha antiga pelo resultado da formula excel


#Mostrar apenas columas 
contatos.loc[contatos.index=='Sem interação', 'Nome do cliente']

#Criar loop
for i, mensagem in enumerate(contatos):
    nome = contatos.loc[i,'Nome do cliente']
    telefone = contatos.loc[i, 'Telefone']

    print(nome)
    print(telefone)




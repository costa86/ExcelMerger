instruc='''Notas:
1.Todos os arquivos de excel nessa pasta serão unificados e organizados por abas de acordo com a coluna selecionada.
2.Eles devem ter extensão (final do arquivo): .xls ou .xlsx.
3.Não há problema se houver arquivos com outras extensões nessa pasta.
------------------------------------------------------------------------- 
'''
print(instruc)
destino = input('Arquivo destino - escolha um nome: ')
import os,time
import pandas as pd
arquivos = []
excel = ('.xls','.xlsx')
for cada in os.listdir():
    if cada.endswith(excel):
        arquivos.append(cada)
frames = []
for i in arquivos:
    a = pd.read_excel(i)
    frames.append(a)
dataFrame = pd.concat(frames)
print('Colunas disponíveis: ')
for i in dataFrame.columns: print (i)
coluna = input('Arquivos de origem - coluna a ser filtrada: ')
abas = dataFrame[coluna].unique()
with pd.ExcelWriter(destino+'.xls') as writer:
        for i in abas:
            dataFrame[dataFrame[coluna]==i].to_excel(writer,index=False,sheet_name=(str(i)))
print('Arquivo {} criado com sucesso!'.format(destino))			
time.sleep(3)

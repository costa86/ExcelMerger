import os,time
import pandas as pd
instruc='''Instructions:
1.All the Excel files (.xls or .xlsx) in this folder will be merged into a new single file organized by sheets according to the chosen columm
2.Any other files with different extensions will be ignored, so don't worry about them.
------------------------------------------------------------------------- 
'''
print(instruc)
destino = input('Output file - choose a name: ')
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
print('Available columms: ')
for i in dataFrame.columns: 
    print (i)
coluna = input('Input file - pick a columm to generate sheets from: ')
abas = dataFrame[coluna].unique()
with pd.ExcelWriter(destino+'.xls') as writer:
        for i in abas:
            dataFrame[dataFrame[coluna]==i].to_excel(writer,index=False,sheet_name=(str(i)))
print(f'File {destino} created successfully!')			
time.sleep(3)

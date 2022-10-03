import pandas as pd
import numpy as np
import math


df = pd.read_excel('visitas_tecnicas.xlsx', usecols='B, C, F, H, I, J')
array = df.values
output = []

i = 0 
for x in array:
    try:
        aux = x[0].split(',')
        if len(aux) <= 1:
            aux.append(0)
        if str(x[2]) == 'nan':
            x[2] = 'Nao Atribuido'
        if type(aux[1]) == str:    
            aux[1] = aux[1].strip(' ').strip('"')
        if aux[1] == '' or aux[1] == ' ':
            aux[1] = '0'
        try:
            if aux[1][0:1].isnumeric():
                aux[1] = aux[1][0:1]
            if aux[1][0:2].isnumeric():
                aux[1] = aux[1][0:2]
        except:
            aux[1] = '0'
        if aux[0] != 'MOBS' and x[1] != 'TRIAGEM DE EQUIPAMENTOS':
            output.append([aux[0], aux[1], x[1], x[2], x[3], x[4], x[5]])
        i += 1
    except:
        print("error", x[0])


i = 0
final_output = output[:]
for x in output:
    n = x[3].count(';')
    names = x[3].split(';')
    if n >= 0:
        final_output[i].insert(0, 1/(n+1))
        try:    
            final_output[i].append(1/(n+1) * float(final_output[i][2]))
        except:
            final_output[i].append(0)
    if n >= 1:
        aux = final_output[i][:]
        for j in range(1, n+1):
            final_output.insert(i, aux[:])
    i += 1 + n
    

for i in range(len(final_output)):
    n = final_output[i][4].count(';')
    names = final_output[i][4].split(';')
    for x in names:
        final_output[i][4] = x
        i += 1

#for x in final_output:
    #print(x)

header = ['PONDERAMENTO', 'CLIENTE', 'QTD', 'TAREFA', 'AGENTE', 'Criado em', 'Data de inicio', 'Data de conclusao', 'ponderacao ativos']

#final_output.insert(0, titulo)

df = pd.DataFrame(final_output)
print(df)

with pd.ExcelWriter('final.xlsx',
                    mode='w') as writer: 
    df.to_excel(writer, index=False, header=header, sheet_name='Instalacoes')
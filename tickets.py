import pandas as pd
import numpy as np
import math

df_open = pd.read_excel('abertos.xlsx', usecols='A,D,E,F,L,N')
array_open = df_open.values
output_open = []

for x in array_open:
    if x[4] == 'TELEMETRIA (Sistema M2)':
        x[4] = 'M2'
    if str(x[5]) == 'nan':
        x[5] = 'Nao Atribuido'
    if str(x[4]) == 'nan':
        x[4] = 'Nao Atribuido'
    output_open.append([x[4], x[1], x[2], x[5], x[0], x[3]])

header = ['PRODUTO', 'Agente', 'Grupo', 'Cliente', 'n do ticket', 'data de abertura']
df = pd.DataFrame(output_open)

with pd.ExcelWriter('final.xlsx', if_sheet_exists='replace',
                    mode='a') as writer:  
    df.to_excel(writer, sheet_name='Tickets', index=False, header=header)


df_closed = pd.read_excel('fechados.xlsx', usecols='A, D, E, F, G, I, L, N')
array_closed = df_closed.values
output_closed = []

#produto agente grupo cliente numeroTick DataFecha DataAber PrimeiraResp
for x in array_closed:
    if str(x[6]) == 'nan':
        x[6] = 'Nao Atribuido'
    if str(x[1]) == 'nan':
        x[1] = 'Nao Atribuido'
    if str(x[2]) == 'nan':
        x[2] = 'Nao Atribuido'
    if str(x[7]) == 'nan':
        x[7] = 'Nao Atribuido'
    if x[6] == 'TELEMETRIA (Sistema M2)':
        x[6] = 'M2'
    output_closed.append([x[6], x[1], x[2], x[7], x[0], x[4], '', x[3], '', x[5]])

df = pd.DataFrame(output_closed)
header = ['PRODUTO', 'Agente', 'Grupo', 'Cliente', 'n do ticket', 'data de Fechamento', 'Fechado a', 'data de Abertura', 'Tempo para solucao', 'resposta']

with pd.ExcelWriter('final.xlsx', if_sheet_exists='replace',
                    mode='a') as writer:  
    df.to_excel(writer, sheet_name='Tickets Fechados', index=False, header=header)

from csv import excel
import PySimpleGUI as sg
import pandas as pd

xlsx = 'data.xlsx'
df_xlsx = pd.read_excel(xlsx)


layout = [
    [sg.Text('Preencha os campos:')],
    [sg.Text('Nick do LoL', size = (10,1)), sg.InputText(key = 'Nick')],
    [sg.Text('ELO', size = (10,1)), sg.Combo(['Ferro', 'Bronze', 'Prata', 'Ouro', 'Platina', 'Diamante', 
    'Mestre', 'Grão-Mestre', 'Desafiante'], key = 'Liga'), sg.Combo(['I', 'II', 'III', 'IV'], key = 'Divisão')],
    [sg.Text('Roles', size = (10,1)), sg.Combo(['Top', 'Jungler', 'Mid', 'ADC', 'Suporte'], key = 'Primária'), 
    sg.Combo(['Top', 'Jungler', 'Mid', 'ADC', 'Suporte'], key = 'Secundária')],
    [sg.Submit(), sg.Exit()]
]

window = sg.Window('Exportar para Excel', layout)

while True:
    evento, valores = window.read()
    if evento == sg.WIN_CLOSED or evento == 'Exit':
        break
    if evento == 'Submit':
        df_xlsx = df_xlsx.append(valores, ignore_index = True)
        df_xlsx.to_excel(xlsx, index = False)
        sg.popup('Salvo')
window.close()
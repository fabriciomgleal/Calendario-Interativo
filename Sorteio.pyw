import PySimpleGUI as sg
import calendar
import datetime
import pandas

sorteios = pandas.read_excel('C:/Users/Fabrício/Documents/GitHub/Calendario-Interativo/Calendario-Interativo/Teste XLSX Sorteios.xlsx', sheet_name='Planilha1')

produtos_primeira_quarta_todos = sorteios.loc[sorteios['Data']==('Primeira Quarta')]

produtos_primeira_quarta_mai_nov = sorteios.loc[sorteios['Data']==('Primeira Quarta Mai Nov')]

produtos_primeira_quarta_dez = sorteios.loc[sorteios['Data']==('Primeira Quarta Dez')]

produtos_primeiro_sabado_todos = sorteios.loc[sorteios['Data']==('Primeiro Sabado')]

primeira_quarta = []
for month in range (1,13):
	mycal = calendar.monthcalendar(datetime.datetime.now().year,month)
	week1 = mycal[0]
	week2 = mycal[1]
	if week1[calendar.WEDNESDAY] != 0:
		primeiraquarta = week1[calendar.WEDNESDAY]
	else:
		primeiraquarta = week2[calendar.WEDNESDAY]
	primeira_quarta.append((str(primeiraquarta)+'/'+str(month)))

primeiro_sabado = []
for month in range (1,13):
	mycal = calendar.monthcalendar(datetime.datetime.now().year,month)
	week1 = mycal[0]
	week2 = mycal[1]
	if week1[calendar.SATURDAY] != 0:
		primeirosabado = week1[calendar.SATURDAY]
	else:
		primeirosabado = week2[calendar.SATURDAY]
	primeiro_sabado.append((str(primeirosabado)+'/'+str(month)))

sg.theme('DarkGreen4')
layout = [[sg.CalendarButton('Calendario', key='dia', format='%d/%m', default_date_m_d_y=(1, None, datetime.datetime.now().year))],
            [sg.Button('Checar data'),sg.Exit()]]

janela = sg.Window('Calendario', layout, no_titlebar=False, size=(150,80))

while True:
    evento, valores=janela.read()
    dia = valores['dia']
    if evento in (sg.WIN_CLOSED, 'Exit'):
        break
    elif evento=='Checar data':
        data = valores['dia'].replace('0','')
        mes = data.split('/')[1]
        if data in primeira_quarta:
            if mes in ('1','2','3','4','6','7','8','9','10'):
                sg.popup(dia, produtos_primeira_quarta_todos[['Nome','Modalidade','Maximo de Sorteio']], title='Produtos primeira quarta')
            elif mes in ('5','11'):
                sg.popup(dia, produtos_primeira_quarta_todos[['Nome','Modalidade','Maximo de Sorteio']], produtos_primeira_quarta_mai_nov[['Nome','Modalidade','Maximo de Sorteio']], title='Produtos primeira quarta')
            elif mes in ('12'):
                sg.popup(dia, produtos_primeira_quarta_todos[['Nome','Modalidade','Maximo de Sorteio']], produtos_primeira_quarta_dez[['Nome','Modalidade','Maximo de Sorteio']], title='Produtos primeira quarta')
        elif data in primeiro_sabado:
            sg.popup(dia, produtos_primeiro_sabado_todos[['Nome','Modalidade','Maximo de Sorteio']], title='Produtos primeiro sabado')
        else:
            sg.popup('Selecione outra data.')    
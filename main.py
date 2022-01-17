import PySimpleGUI as sg
import pandas as pd

# Add some color to the window
sg.theme('Dark')

EXCEL_FILE = 'Data_Entry.xlsx'
df = pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Name', size=(15,1)), sg.InputText(key='Name')],
    [sg.Text('Phone Number', size=(15,1)), sg.InputText(key='Phone')],
    [sg.Text('Email', size=(15,1)), sg.InputText(key='Email')],
    [sg.Text('Gender', size=(15,1)), sg.Combo(['Female', 'Male', 'Others'], key='Gender')],
    [sg.Text('Nationality', size=(15,1)),
                            sg.Checkbox('Indian', key='Indian'),
                            sg.Checkbox('Foreign', key='Foreign')],
    [sg.Text('No. of Children', size=(15,1)), sg.Spin([i for i in range(0,16)],
                                                       initial_value=0, key='Children')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('Kids Store - Data Collection', layout)

def clear_input():
    for key in values:
        window[key]('')
    return None

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        df = df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data saved!')
        clear_input()
window.close()
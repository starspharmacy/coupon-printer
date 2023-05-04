import PySimpleGUI as sg
import win32api
import win32print
import os
import openpyxl
from openpyxl.utils import get_column_letter
import pythoncom
import pywintypes
from win32com.client import DispatchEx


all_printers = [printer[2] for printer in win32print.EnumPrinters(2)]

sg.theme('DarkBlue')
sg.set_options(font=('Helvetica', 11))

layout = [
    [sg.Text('Number of copies to print:'), sg.Input(default_text='1', size=(10, 1), key='-NUM_COPIES-', focus=True)],
    [sg.Text('Date:'), sg.Input(key='-DATE-', size=(20, 1)), sg.Text('Customer Mobile Number:'), sg.Input(key='-MOBILE-', size=(20, 1))],
    [sg.Text('Select printer:'), sg.DropDown(all_printers, default_value=win32print.GetDefaultPrinter(), size=(30, 1), key='-PRINTER-')],
    [
        sg.Button('PRINT', size=(10, 1)), 
        sg.Button('SAVE', size=(10, 1)), 
        sg.Button('CLOSE', size=(10, 1), button_color=('white', 'red'))
    ]
]

window = sg.Window('Stars Coupon - (c)HYBRID', layout)

while True:
    event, values = window.read()
    
    if event == sg.WINDOW_CLOSED or event == 'CLOSE':
        break
    
    if event == 'PRINT' or (event == '-NUM_COPIES-' and values[event] and values[event][-1] == '\n'):
        file_path = os.path.join(os.path.dirname(__file__), 'coupon.xlsx')
        printer_name = values['-PRINTER-']
        try:
            num_copies = int(values['-NUM_COPIES-'])
        except ValueError:
            sg.popup('Please enter a valid number of copies.')
            continue
        
        if values['-DATE-'] or values['-MOBILE-']:
            pythoncom.CoInitialize()
            excel = DispatchEx('Excel.Application')
            excel.Visible = False
            workbook = excel.Workbooks.Open(file_path)
            worksheet = workbook.ActiveSheet
            
            if values['-DATE-']:
                worksheet.Range('E1').Value = values['-DATE-']
            if values['-MOBILE-']:
                worksheet.Range('B1').Value = values['-MOBILE-']
                
            workbook.Save()
            workbook.Close()
            excel.Quit()
        
        for i in range(num_copies):
            win32api.ShellExecute(
                0,
                'print',
                file_path,
                f'/d:"{printer_name}"',
                '.',
                0
            )
            
    if event == 'SAVE':
        file_path = os.path.join(os.path.dirname(__file__), 'coupon.xlsx')
        if not (values['-DATE-'] or values['-MOBILE-']):
            sg.popup('Please enter a date or customer mobile number.')
            continue
        
        pythoncom.CoInitialize()
        excel = DispatchEx('Excel.Application')
        excel.Visible = False
        workbook = excel.Workbooks.Open(file_path)
        worksheet = workbook.ActiveSheet
        
        if values['-DATE-']:
            worksheet.Range('E1').Value = values['-DATE-']
        if values['-MOBILE-']:
            worksheet.Range('B1').Value = values['-MOBILE-']
            
        workbook.Save()
        workbook.Close()
        excel.Quit()
        
    if event == 'UPDATE':
        check_for_updates()

window.close()

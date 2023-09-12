import pandas as pd
import argparse
import sys
import PySimpleGUI as sg

class ValuesToClass(object):
    def __init__(self,values):
        for key in values:
            setattr(self, key, values[key])

if __name__ == '__main__':
    if len(sys.argv) != 1:
        argParser = argparse.ArgumentParser(description='Filter all sheets in an excel file based on a given column and value set.')
        argParser.add_argument('-i', '--input_excel', action='store', required=True)
        argParser.add_argument('-f', '--filter_col', action='store', required=True)
        argParser.add_argument('-v', '--filter_vals', action='store', required=True, type=lambda path: pd.read_excel(path))
        args = argParser.parse_args()

    else:
        sg.theme('Dark Blue 17')

        layout = [[sg.Text('Excel File'), sg.Input(key='input_excel'), sg.FolderBrowse()],
                  [sg.Text('Filtering Column'), sg.Input(key='filter_col')],
                  [sg.Text('Filtering Values'), sg.Input(key='filter_vals')]
                  [sg.Submit(), sg.Cancel()]]
        
        window = sg.Window('CSV Conversion Script', layout)

        event,  values = window.read()
        window.close()

        if event=='Cancel':
            quit()
        else:
            args = ValuesToClass(values)
        

    input_dfs = pd.read_excel(args.input_excel, sheet_name=None)
    filter_vals = args.filter_vals[args.filter_col].astype(str).unique()
    with pd.ExcelWriter('{}_filtered.xlsx'.format(args.input_excel.split('.xls')[0])) as writer:
        for sname, df in input_dfs.items():
            try:
                df = df[df[args.filter_col].astype(str).isin(filter_vals)]
                df.to_excel(writer, sheet_name=sname, index=False, na_rep='N/A')
            except:
                print(sname)
                continue

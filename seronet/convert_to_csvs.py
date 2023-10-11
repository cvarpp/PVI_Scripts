import pandas as pd
import sys
import glob
import os
import argparse
import PySimpleGUI as sg

class ValuesToClass(object):
    def __init__(self,values):
        for key in values:
            setattr(self, key, values[key])

if __name__ == '__main__':
    if len(sys.argv) != 1:
        argParser = argparse.ArgumentParser(description='Read Excel files in input_dir and create corresponding csv files in output_dir')
        argParser.add_argument('-i', '--input_dir', action='store', required=True)
        argParser.add_argument('-o', '--output_dir', action='store', required=True)
        args = argParser.parse_args()

    else:
        sg.theme('Dark Blue 17')

        layout = [[sg.Text('Input Folder'), sg.Input(key='input_dir'), sg.FolderBrowse()],
                  [sg.Text('Output Folder'), sg.Input(key='output_dir'), sg.FolderBrowse()],
                  [sg.Submit(), sg.Cancel()]]
        
        window = sg.Window('CSV Conversion Script', layout)

        event,  values = window.read()
        window.close()

        if event=='Cancel':
            quit()
        else:
            args = ValuesToClass(values)
    for fname in glob.glob('{}/*xls*'.format(args.input_dir)):
        print("Converting", fname)
        pd.read_excel(fname, na_filter=False, keep_default_na=False).to_csv('{}/{}.csv'.format(args.output_dir, fname.split(os.sep)[-1].split('.')[0]), index=False)
        print(fname, "converted!")
    print("Done!")

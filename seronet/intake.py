import pandas as pd
import numpy as np
import datetime
from dateutil.relativedelta import relativedelta
import argparse
import util
import sys
import PySimpleGUI as sg
from seronet.accrual import accrue
import os
from helpers import ValuesToClass
import warnings

def make_intake(df_accrual, ecrabs, dfs_clin, seronet_key):
    pass

if __name__ == '__main__':
    if len(sys.argv) != 1:
        argParser = argparse.ArgumentParser(description='Make files for monthly data submission.')
        argParser.add_argument('-c', '--use_cache', action='store_true')
        argParser.add_argument('-s', '--report_start', action='store', type=pd.to_datetime)
        argParser.add_argument('-e', '--report_end', action='store', type=pd.to_datetime)
        argParser.add_argument('-d', '--debug', action='store_true')
        args = argParser.parse_args()
    else:
        sg.theme('Dark Blue 17')

        layout = [[sg.Text('Accrual')],
                  [sg.Checkbox("Use Cache", key='use_cache', default=False), \
                    sg.Checkbox("debug", key='debug', default=False)],
                  [sg.Text('Start date'), sg.Input(key='report_start', default_text='1/1/2021'), sg.CalendarButton(button_text="choose date",close_when_date_chosen=True, target="report_start", format='%m/%d/%Y')],
                        [sg.Text('End date'), sg.Input(key='report_end', default_text='12/31/2025'), sg.CalendarButton(button_text="choose date",close_when_date_chosen=True, target="report_end", format='%m/%d/%Y')],
                    [sg.Submit(), sg.Cancel()]]

        window = sg.Window("Accrual Generation Script", layout)

        event, values = window.read()
        window.close()

        if event =='Cancel':
            quit()
        else:
            values['report_start'] = pd.to_datetime(values['report_start'])
            values['report_end'] = pd.to_datetime(values['report_start'])
            args = ValuesToClass(values)
    df_accrual, ecrabs, dfs_clin, seronet_key = accrue(args)
    make_intake(df_accrual, ecrabs, dfs_clin, seronet_key)
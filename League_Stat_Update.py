## Updating League_Stat.xlsx using Data from Updated League_Result.xlsx

import pandas as pd
import numpy as np
import xlwings as xw

## Reading League Result as Data for Update
locale = 'Raymond'
df_res  = pd.ExcelFile('C:/Users/'+locale+'/Google Drive/Python/FPL/League_Result.xlsx').parse()
df_stat = pd.ExcelFile('C:/Users/'+locale+'/Google Drive/Python/FPL/League_Stat.xlsx').parse()
team_sequence = list(df_stat['Team Name'])

## Set up League Result from xwings to update values
book = xw.Book('C:/Users/'+locale+'/Google Drive/Python/FPL/League_Stat.xlsx')
sheet = book.sheets['Sheet1']

## Updating Sequentially
for [i,team] in enumerate(team_sequence):
    gs_home = np.sum(df_res['Home Score'][df_res['Home'] == team])/len(df_res[df_res['Home'] == team])
    gs_away = np.sum(df_res['Away Score'][df_res['Away'] == team])/len(df_res[df_res['Away'] == team])
    gs_totl = np.sum(df_res['Home Score'][df_res['Home'] == team]) + np.sum(df_res['Away Score'][df_res['Away'] == team])
    gc_home = np.sum(df_res['Away Score'][df_res['Home'] == team])/len(df_res[df_res['Home'] == team])
    gc_away = np.sum(df_res['Home Score'][df_res['Away'] == team])/len(df_res[df_res['Away'] == team])
    gc_totl = np.sum(df_res['Away Score'][df_res['Home'] == team]) + np.sum(df_res['Home Score'][df_res['Away'] == team])
    sheet.range('H' + str(i + 2)).value = gs_totl
    sheet.range('J' + str(i + 2)).value = gs_home
    sheet.range('K' + str(i + 2)).value = gs_away
    sheet.range('I' + str(i + 2)).value = gs_totl
    sheet.range('L' + str(i + 2)).value = gc_home
    sheet.range('M' + str(i + 2)).value = gc_away
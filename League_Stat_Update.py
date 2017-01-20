## Updating League_Stat.xlsx using Data from Updated League_Result.xlsx

import pandas as pd
import numpy as np
import xlwings as xw

# locale = 'Raymond'
locale = 'santoray'

def update_league_stat():
    ## Reading League Result as Data for Update
    df_res  = pd.ExcelFile('C:/Users/'+locale+'/Google Drive/Python/FPL/League_Result.xlsx').parse()
    df_stat = pd.ExcelFile('C:/Users/'+locale+'/Google Drive/Python/FPL/League_Stat.xlsx').parse()
    team_sequence = list(df_stat['Team Name'])

    ## Set up League Result from xwings to update values
    book = xw.Book('C:/Users/'+locale+'/Google Drive/Python/FPL/League_Stat.xlsx')
    sheet = book.sheets['Stat']

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
        sheet.range('I' + str(i + 2)).value = gc_totl
        sheet.range('L' + str(i + 2)).value = gc_home
        sheet.range('M' + str(i + 2)).value = gc_away

def update_fixture_predictions():
    ## Reading League Result as Data for Update
    df_res  = pd.ExcelFile('C:/Users/' + locale + '/Google Drive/Python/FPL/League_Result.xlsx').parse()
    # df_fxt = pd.ExcelFile('C:/Users/'+locale+'/Google Drive/Python/FPL/Fixture_Predictions2.xlsx').parse()
    df_stat = pd.ExcelFile('C:/Users/' + locale + '/Google Drive/Python/FPL/League_Stat.xlsx').parse()
    team_sequence = list(df_stat['Team Name'])
    latest_gw = max(np.array(df_res['GW']))

    ## Set up League Result from xwings to update values
    book = xw.Book('C:/Users/' + locale + '/Google Drive/Python/FPL/Fixture_Predictions2.xlsx')
    poisson_dict = {}

    ## Updating Sequentially
    for team in team_sequence:
        sheet = book.sheets[team]
        gs_home_estimate = []
        gs_away_estimate = []
        gc_home_estimate = []
        gc_away_estimate = []
        for j in range(1,latest_gw+1):
            df_gw = df_res[['GW','Home','Away']][(df_res['Home'] == team) | (df_res['Away'] == team)]
            df_gw.index = range(1,latest_gw+1)
            team_home = list(df_gw['Home'][df_gw['GW'] == j])[0]
            team_away = list(df_gw['Away'][df_gw['GW'] == j])[0]
            if team == team_home:
                gs = int(df_res['Home Score'][(df_res['GW'] == j) & (df_res['Home'] == team)])
                gc = int(df_res['Away Score'][(df_res['GW'] == j) & (df_res['Home'] == team)])
                gs_home_estimate.append(gs)
                gc_home_estimate.append(gc)
                sheet.range('D'+str(j+1)).value = 'H'
                sheet.range('I'+str(j+1)).value = team_away
            elif team == team_away:
                gs = int(df_res['Away Score'][(df_res['GW'] == j) & (df_res['Away'] == team)])
                gc = int(df_res['Home Score'][(df_res['GW'] == j) & (df_res['Away'] == team)])
                gs_away_estimate.append(gs)
                gc_away_estimate.append(gc)
                sheet.range('D' + str(j + 1)).value = 'A'
                sheet.range('I' + str(j + 1)).value = team_home
            else:
                print('ERROR!!')
                break
            sheet.range('E' + str(j + 1)).value = gc
            sheet.range('G' + str(j + 1)).value = gs
        poisson_dict[team] = [np.mean(gs_home_estimate), np.mean(gs_away_estimate), np.mean(gc_home_estimate), np.mean(gc_away_estimate)]

    ## Updating Projection based on Poisson Average
    for team in team_sequence:
        sheet = book.sheets[team]
        for k in range(1,39):
            oppt = sheet.range('C'+str(k+1)).value
            ha_status = sheet.range('D'+str(k+1)).value
            print(k, team, oppt, poisson_dict[team][3], poisson_dict[oppt][0], poisson_dict[team][1], poisson_dict[oppt][2])
            if ha_status== 'H':
                sheet.range('F'+str(k+1)).value = poisson_dict[team][2]*poisson_dict[oppt][1]
                sheet.range('H'+str(k+1)).value = poisson_dict[team][0]*poisson_dict[oppt][3]
            elif ha_status == 'A':
                sheet.range('F'+str(k+1)).value = poisson_dict[team][3]*poisson_dict[oppt][0]
                sheet.range('H'+str(k+1)).value = poisson_dict[team][1]*poisson_dict[oppt][2]
            else:
                print('ERROR2!!')
                break

    ## Output All team Poisson projection
    for team in team_sequence:
        print(team, poisson_dict[team])

# update_league_stat()
update_fixture_predictions()
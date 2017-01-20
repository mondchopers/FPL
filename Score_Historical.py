import pandas as pd
import numpy as np

# Team Abbreviation Dictionnary
team_abr = {'ARS': 'Arsenal', 'BOU': 'Bournemouth', 'BUR': 'Burnley', 'CHE': 'Chelsea', 'CRY': 'CrystalPalace',
            'EVE': 'Everton', 'HUL': 'HullCity', 'LEI': 'Leicester', 'LIV': 'Liverpool', 'MCI': 'ManchesterCity',
            'MUN': 'ManchesterUnited', 'MID': 'Middlesbrough', 'SOU': 'Southampton', 'STK': 'StokeCity',
            'SUN': 'Sunderland', 'SWA': 'Swansea', 'TOT': 'Tottenham', 'WAT': 'Watford', 'WBA': 'WestBrom',
            'WHU': 'WestHam'}
unique_team = ['Arsenal','Bournemouth','Burnley','Chelsea','CrystalPalace','Everton','HullCity','Leicester','Liverpool',
               'ManchesterCity','ManchesterUnited','Middlesbrough','Southampton','StokeCity','Sunderland','Swansea',
               'Tottenham','Watford','WestBrom','WestHam']
unique_id = ['ARS','BOU','BUR','CHE','CRY','EVE','HUL','LEI','LIV','MCI','MUN','MID','SOU','STK','SUN','SWA','TOT',
             'WAT','WBA','WHU']

# locale = 'Raymond'
locale = 'santoray'
threshold = 5

## Load Scoring Data Frame
last_gw = 21
folder    = 'C:/Users/'+locale+'/Google Drive/Python/FPL/'
history_dict = {}
for i in range(5,last_gw+1):
    history_dict[i] = pd.read_csv(folder + 'FPL16-GW' + str(i) + '.csv', engine='python')
    history_dict[i].index = range(0,len(history_dict[i]))
    cdnm = []
    for jj in range(0,len(history_dict[i])):
        cdnm.append(str(history_dict[i]['Surname'][jj])+'@'+str(history_dict[i]['FirstName'][jj]))
    history_dict[i]['Codename'] = cdnm
    history_dict[i].sort_values(by='Codename')
    history_dict[i].index = range(0,len(history_dict[i]))

## Filtering player with criteria
## Must exist between GW 5 - GW latest (does't have to be in all)
## Must still exist at GW latest
## Must have points of > threshold
eligible = set(history_dict[5]['Codename']).intersection(set(history_dict[6]['Codename']))
for i in range(7,last_gw):
    eligible = eligible.union(set(history_dict[i]['Codename']))
eligible = eligible.intersection(set(history_dict[last_gw]['Codename'].loc[history_dict[last_gw]['TotalPoints'] > threshold]))
eligible = list(eligible)

## Create a new Data Frame To Store Histoical Scoring
## Append Global Data from Latest GW: Team, Position, Total Points
df_hist = pd.DataFrame()
df_hist['CD'] = history_dict[last_gw]['Codename'].loc[history_dict[last_gw]['Codename'].isin(eligible)]
# df_hist['Team'] = history_dict[last_gw]['Team'].loc[history_dict[last_gw]['Codename'].isin(eligible)]
df_hist['Pos'] = history_dict[last_gw]['PositionsList'].loc[history_dict[last_gw]['Codename'].isin(eligible)]
df_hist['TotalPt'] = history_dict[last_gw]['TotalPoints'].loc[history_dict[last_gw]['Codename'].isin(eligible)]
df_hist['Cost'] = history_dict[last_gw]['Cost'].loc[history_dict[last_gw]['Codename'].isin(eligible)].apply(
    lambda x:x*0.000001)
df_hist = df_hist.sort_values(by='CD')
df_hist.index = range(0,len(df_hist))

## Append Data from Each GW only for eligible players
df_gw_dict = {}
for i in range(5,last_gw+1):
    df_gw = history_dict[i][['Codename','Team','PointsLastRound','MinutesPlayed','GoalsScored','Assists',
                             'GoalsConceded','Saves','Bonus']].loc[history_dict[i]['Codename'].isin(eligible)]
    missing = [x for x in eligible if x not in set(df_gw['Codename'])]
    df_gw.index = range(0,len(df_gw))
    for item in missing:
        df_gw.loc[len(df_gw)+1] = [item, history_dict[last_gw]['Team'][history_dict[last_gw]['Codename'] ==
                                                                       item].tolist()[0],0,0,0,0,0,0,0]
    df_gw = df_gw.sort_values(by='Codename')
    df_gw.index = range(0,len(df_gw))
    df_gw_dict[i] = df_gw

## Append Data that is the difference between current and last GW
# for i in range(5,last_gw+1):
#     df_hist['CD'+str(i)] = df_gw_dict[i]['Codename']
for i in range(5,last_gw+1):
    df_hist['PtGW'+str(i)] = df_gw_dict[i]['PointsLastRound']
for i in range(2,6):
    q = np.array(df_hist['PtGW'+str(last_gw)])
    for j in range(1,i):
        q += np.array(df_hist['PtGW'+str(last_gw-j)])
    df_hist['PtLast'+str(i)] = q
for i in range(6, last_gw + 1):
    df_hist['MinGW' + str(i)] = np.array(df_gw_dict[i]['MinutesPlayed'])-np.array(df_gw_dict[i-1]['MinutesPlayed'])
for i in range(6, last_gw + 1):
    df_hist['GoalGW' + str(i)] = np.array(df_gw_dict[i]['GoalsScored'])-np.array(df_gw_dict[i-1]['GoalsScored'])
for i in range(6, last_gw + 1):
    df_hist['AsstGW' + str(i)] = np.array(df_gw_dict[i]['Assists'])-np.array(df_gw_dict[i-1]['Assists'])
for i in range(6, last_gw + 1):
    df_hist['ConcGW' + str(i)] = np.array(df_gw_dict[i]['GoalsConceded'])-np.array(df_gw_dict[i-1]['GoalsConceded'])
for i in range(6, last_gw + 1):
    df_hist['SaveGW' + str(i)] = np.array(df_gw_dict[i]['Saves'])-np.array(df_gw_dict[i-1]['Saves'])
for i in range(6, last_gw + 1):
    df_hist['BonusGW' + str(i)] = np.array(df_gw_dict[i]['Bonus'])-np.array(df_gw_dict[i-1]['Bonus'])
for i in range(5,last_gw+1):
    df_hist['TeamGW' + str(i)] = [team_abr[y] for y in df_gw_dict[i]['Team'].tolist()]
    # df_hist['TeamGW'+str(i)].replace(unique_id, unique_team, inplace=True)

## Loading Stats and Fixture Data Frame to Append Average Projections
df_stat = pd.ExcelFile('C:/Users/'+locale+'/Google Drive/Python/FPL/League_Stat.xlsx').parse()
df_fixt_dict = {}
for team in unique_team:
    df_fixt_dict[team] = pd.ExcelFile('C:/Users/'+locale+'/Google Drive/Python/FPL/Fixture_Predictions.xlsx').parse(
        sheetname=team)

# ## Append team against info
for i in range(5, last_gw + 1):
    df_hist['OpptGW'+str(i)] = df_hist['TeamGW' + str(i)].apply(lambda x:df_fixt_dict[x]['Oppt'][i-1])
for i in range(5, last_gw + 1):
    df_hist['GSGW'+str(i)] = df_hist['TeamGW' + str(i)].apply(lambda x:df_fixt_dict[x]['Actual Scored'][i-1])
    df_hist['GCGW'+str(i)] = df_hist['TeamGW' + str(i)].apply(lambda x:df_fixt_dict[x]['Actual Conceeded'][i-1])
    df_hist['GSExpGW'+str(i)] = df_hist['TeamGW' + str(i)].apply(lambda x:df_fixt_dict[x]['Expected Goals Scored'][
        i-1])
    df_hist['GCExpGW'+str(i)] = df_hist['TeamGW' + str(i)].apply(lambda x:df_fixt_dict[x]['Expected Goals Scored'][i-1])

## Saving Result
writer = pd.ExcelWriter('C:/Users/'+locale+'/Google Drive/Python/FPL/Score_History.xlsx')
df_hist.to_excel(writer)
writer.save()
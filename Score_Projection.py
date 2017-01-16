import pandas as pd
import numpy as np
from scipy.stats import poisson
from sklearn import linear_model

# Team Abbreviation Dictionnary
team_abr = {'ARS': 'Arsenal', 'BOU': 'Bournemouth', 'BUR': 'Burnley', 'CHE': 'Chelsea', 'CRY': 'CrystalPalace',
            'EVE': 'Everton', 'HUL': 'HullCity', 'LEI': 'Leicester', 'LIV': 'Liverpool', 'MCI': 'ManchesterCity',
            'MUN': 'ManchesterUnited', 'MID': 'Middlesbrough', 'SOU': 'Southampton', 'STK': 'StokeCity',
            'SUN': 'Sunderland', 'SWA': 'Swansea', 'TOT': 'Tottenham', 'WAT': 'Watford', 'WBA': 'WestBrom',
            'WHU': 'WestHam'}

locale = 'Raymond'

def calculate_asstgoal_ratio(df_latest):
    ## Calculating Assist and Goal per minute
    goal_pm = []
    asst_pm = []
    for i in range(0,len(df_latest)):
        # team_goal.append(np.sum(df_latest['GoalsScored'][df_latest['Team'] == df_latest['Team'][i]]))
        # team_asst.append(np.sum(df_latest['Assists'][df_latest['Team'] == df_latest['Team'][i]]))
        if df_latest['MinutesPlayed'][i] > 15*last_gw: ## Minimum average 15 mpg to qualify
            goal_pm.append(float(df_latest['GoalsScored'][i])/float(df_latest['MinutesPlayed'][i]))
            asst_pm.append(float(df_latest['Assists'][i])/float(df_latest['MinutesPlayed'][i]))
        else:
            goal_pm.append(0)
            asst_pm.append(0)
    df_latest['GoalPerMin']  = goal_pm
    df_latest['AsstPerMin']  = asst_pm
    ## Appending Assist and Goal Ratio
    asst_ratio = []
    goal_ratio = []
    for j in range(0, len(df_latest)):
        goal_ratio.append(float(df_latest['GoalPerMin'][j])/np.sum(df_latest['GoalPerMin'][df_latest['Team'] == df_latest['Team'][j]]))
        asst_ratio.append(float(df_latest['AsstPerMin'][j])/np.sum(df_latest['AsstPerMin'][df_latest['Team'] == df_latest['Team'][j]]))
    df_latest['GoalRatio'] = goal_ratio
    df_latest['AsstRatio'] = asst_ratio

## Load Scoring Data Frame
last_gw = 20
gw_ahead = 5
folder    = 'C:/Users/'+locale+'/Google Drive/Python/FPL/'
df_latest = pd.read_csv(folder + 'FPL16-GW' + str(last_gw) + '.csv', engine='python')
df_latest.index = np.arange(0,len(df_latest))

## Load Stats and Fixtures Data Frame
df_stat = pd.ExcelFile('C:/Users/'+locale+'/Google Drive/Python/FPL/League_Stat.xlsx').parse()
df_fixture_dict = {}
for cle in team_abr.keys():
    df_fixture_dict[cle] = pd.ExcelFile('C:/Users/'+locale+'/Google Drive/Python/FPL/Fixture_Predictions.xlsx').parse(sheetname=team_abr[cle])
df_prediction = pd.ExcelFile('C:/Users/'+locale+'/Google Drive/Python/FPL/Prediction.xlsx').parse()

## Appending goal and assist ratio into the spreadsheet
calculate_asstgoal_ratio(df_latest)

## Estimating Saves: Team Reduction factor = opp shot/gm - (goal conceeded + saves)
## Therefore save predictor  = Save Ratio *(Opp shot/gm - Reduction factor) where Save Ratio = Saves/(Saves+Goal Conceeded)
opp_shot_pg = []
# for i in range(0,len(df_latest)):

## Preparing sample
pos = 'DEF'
n_sample = 30
df_latest['Cost_S'] = df_latest['Cost']/1000000
df_sample = df_latest[['FirstName','Surname','PositionsList','Team','TotalPoints','AveragePoints','MinutesPlayed','GoalsScored','Assists','GoalsConceded','Saves','CleanSheets','Bonus','PenaltiesSaved','PenaltiesMissed','YellowCards','RedCards','OwnGoals','GoalRatio','AsstRatio']][df_latest['PositionsList'] == pos].sort(columns='AveragePoints',ascending=[0]).copy()
df_sample.index = np.arange(0,len(df_sample))
df_sample = df_sample[0:n_sample] ## Take Top n_sample scorers as sample
# Preparing Payout Arrays
if pos == 'GLK':
    payout_cscc = np.array([4, 0, -1, -1, -2, -2, -3])
    payout_goal = 6 * np.arange(0, 7, 1)
elif pos == 'DEF':
    payout_cscc = np.array([4, 0, -1, -1, -2, -2, -3])
    payout_goal = 6 * np.arange(0, 7, 1)
elif pos == 'MID':
    payout_cscc = np.array([1, 0, 0, 0, 0, 0, 0])
    payout_goal = 5 * np.arange(0, 7, 1)
elif pos == 'FWD':
    payout_cscc = np.array([0, 0, 0, 0, 0, 0, 0])
    payout_goal = 4 * np.arange(0, 7, 1)
else:
    print('ERROR!!')
payout_asst = 3 * np.arange(0, 7, 1)

score_gw_goal = []
score_gw_asst = []
score_gw_save = []
score_gw_cscc = []
score_gw_bnus = []
score_gw_penr = []
score_gw_fups = []
for z in range(1,1+gw_ahead):
    ## Appending Next Opponent and Home-Away Status
    oppt = []
    ha_status = []
    for ii in range(0,len(df_sample)):
        opp_dict = df_fixture_dict[df_sample['Team'][ii]]
        nextopp = opp_dict['Oppt'][gw_ahead+z-1]
        ha = opp_dict['AtHome'][gw_ahead+z-1]
        oppt.append(nextopp)
        ha_status.append(ha)
    df_sample['Oppt'+str(z)] = oppt
    df_sample['HA'+str(z)] = ha_status

    ## Predicting Goal Amount
    proba_score = []
    proba_conceed = []
    for jj in range(0,len(df_sample)):
        sc = []
        cc = []
        my_team = team_abr[df_sample['Team'][jj]]
        op_team = df_sample['Oppt'+str(z)][jj]
        ha_team = df_sample['HA'+str(z)][jj]
        # print(jj, my_team, op_team, ha_team)
        if ha_team == 'H': ## Team playing at home
            sc.append(float(df_prediction['PoissonMoyenne'][
                (df_prediction['Team'] == op_team) & (df_prediction['Oppt'] == my_team) & (df_prediction['AtHome'] == False)]))
            cc.append(float(df_prediction['PoissonMoyenne'][(df_prediction['Team'] == my_team)
                            & (df_prediction['Oppt'] == op_team) & (df_prediction['AtHome'] == True)]))
        else:
            sc.append(float(df_prediction['PoissonMoyenne'][(df_prediction['Team'] == op_team)
                            & (df_prediction['Oppt'] == my_team) & (df_prediction['AtHome'] == True)]))
            cc.append(float(df_prediction['PoissonMoyenne'][(df_prediction['Team'] == my_team)
                            & (df_prediction['Oppt'] == op_team) & (df_prediction['AtHome'] == False)]))
        proba_score.append(sc)
        proba_conceed.append(cc)
    # df_sample['ScoreProb'] = proba_score
    # df_sample['ConceedProb'] = proba_conceed

    ##Predicting Scores##
    score_cscc = []
    score_goal = []
    score_asst = []
    for kk in range(0,len(df_sample)):
        score_cscc.append(np.sum(poisson.pmf(np.arange(0,7,1), float(proba_conceed[kk][0]))*payout_cscc))
        score_goal.append(np.sum(poisson.pmf(np.arange(0,7,1), float(proba_score[kk][0]))*
                                 payout_goal*df_sample['GoalRatio'][kk]))
        score_asst.append(np.sum(poisson.pmf(np.arange(0, 7, 1), float(proba_score[kk][0])) *
                                 payout_asst*df_sample['AsstRatio'][kk]))

    ## Appending GW Score Factor ##
    score_gw_goal.append(score_goal)
    score_gw_asst.append(score_asst)
    # score_gw_save.append(score_goal)
    score_gw_cscc.append(score_cscc)
    # score_gw_bnus.append(score_goal)
    # score_gw_penr.append(score_goal)
    # score_gw_fups.append(score_goal)

for y in range(1,1+gw_ahead):
    df_sample['GoalSCR'+str(y)] = score_gw_goal[y-1]
    df_sample['AsstSCR' + str(y)] = score_gw_asst[y-1]
    df_sample['CsccSCR'+str(y)] = score_gw_cscc[y-1]
for x in range(1,1+gw_ahead):
    # df_sample['TotalSCR'+str(x)] = df_sample['GoalSCR'+str(y)]
    df_sample['TotalSCR'+str(x)] = np.array(score_gw_goal[x-1]) + np.array(score_gw_asst[x-1]) + np.array(score_gw_cscc[x-1])

## Saving Final Results
writer = pd.ExcelWriter('C:/Users/'+locale+'/Google Drive/Python/FPL/Projection.xlsx')
df_sample.to_excel(writer)
# df_latest.to_excel(writer)
writer.save()



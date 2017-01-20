import xlwings as xw
import pandas as pd

locale = 'Raymond'

team_abr = {'ARS': 'Arsenal', 'BOU': 'Bournemouth', 'BUR': 'Burnley', 'CHE': 'Chelsea', 'CRY': 'CrystalPalace',
            'EVE': 'Everton', 'HUL': 'HullCity', 'LEI': 'Leicester', 'LIV': 'Liverpool', 'MCI': 'ManchesterCity',
            'MUN': 'ManchesterUnited', 'MID': 'Middlesbrough', 'SOU': 'Southampton', 'STK': 'StokeCity',
            'SUN': 'Sunderland', 'SWA': 'Swansea', 'TOT': 'Tottenham', 'WAT': 'Watford', 'WBA': 'WestBrom',
            'WHU': 'WestHam'}

wb = xw.Book('C:/Users/'+locale+'/Google Drive/Python/FPL/League_Result_Backup.xlsx')  # connect to an existing file in the current working directory
sht = wb.sheets('Results')
print(sht.range('A1').value)
# df_fp_last = pd.ExcelFile('C:/Users/'+locale+'/Google Drive/Python/FPL/League_Stat.xlsx').parse()
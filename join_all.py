from datetime import datetime
import pandas as pd
import os

data = str(datetime.now().date())
path = ''
al010 = ''
a = pd.read_excel(al010, sheet_name='Data')
q = []
for i in os.listdir(path):
    if i != 'AL010.xlsm':
        q.append(path + f'/{i}')
for j in q:
    j = pd.read_excel(j, sheet_name='Data')
    a = a.merge(j, how='left')
    a.to_excel(fr'\\RETRU643-NT0001\Common_A\All Rostokino\GADD\ALL.xlsx', startrow=0, index=False)


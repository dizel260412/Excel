import time
from datetime import datetime
import pandas as pd
import os

data = str(datetime.now().date())
path = ''
al010 = ''
a = pd.read_excel(al010, sheet_name='Data')

for i in os.listdir(path):
    t = path + f'/{i}'
    try:
        j = pd.read_excel(t, sheet_name='Data')
        if j['ARTNO'].dtype == "int64":
            a = a.merge(j, how='left')
            a.to_excel(fr'\\RETRU643-NT0001\Common_A\All Rostokino\ALL.xlsx', startrow=0, index=False)
            print(f"Файл {i} добавлен")
            time.sleep(5)
    except Exception as ex:
        print(t)
        continue


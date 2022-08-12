from datetime import datetime
import pandas as pd


file = 'AL010.xlsm'
file_4 = 'SG010.xlsm'
data = str(datetime.now().date())
a = pd.read_excel(file, sheet_name='Data', usecols=['ARTNO', 'ARTNAME_UNICODE'])
e = pd.read_excel(file_4, sheet_name='Data', usecols=['STORAGE_UNICODE', 'SGFLOCATION'])

a['ARTNO'] = a['ARTNO'].astype(str)  ###### Переформатируем данные в формат Object
e['STORAGE_UNICODE'] = e['STORAGE_UNICODE'].astype(str)
i = a.merge(e, how='left', left_on=['ARTNO'], right_on=['STORAGE_UNICODE'])
i['ARTNO'] = i['ARTNO'].astype(int)
i = i[['ARTNO', 'ARTNAME_UNICODE', 'SGFLOCATION']]
i.to_excel('new5.xlsx', index="")  ######Создаем новый файл и запихиваем в него полученные данные
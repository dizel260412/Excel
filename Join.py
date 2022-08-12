from datetime import datetime
import pandas as pd

file = 'AL010.xlsm'
file_1 = 'TT520b.xlsm'
file_2 = 'SLM0003.xlsm'
file_3 = 'AL065.xlsm'
file_4 = 'AL065.xlsm'

data = str(datetime.now().date())
a = pd.read_excel(file, sheet_name='Data')
# b = pd.read_excel(file_1, sheet_name='Data')
# c = pd.read_excel(file_2, sheet_name='Data')
d = pd.read_excel(file_3, sheet_name='Data')

# a = a[['ARTNO', 'ARTNAME_UNICODE']]
# b = b[['ARTNO', 'TRANSTYPE', 'QTY', 'DATE']]
# c = c[['ARTNO', 'SLID']]
# d = d[['ARTNO', 'HFB']]

i = a.merge(d, how='left')  ##########   how='outer'
# i = i[['HFB', 'ARTNO', 'ARTNAME_UNICODE', 'SLID', 'DATE', 'TRANSTYPE', 'QTY']]    ###### ставим колонки на необходимые места

i.to_excel(f'new {data}.xlsx', startrow=0, index=False)   ######Создаем новый файл и запихиваем в него полученные данные

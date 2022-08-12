import pandas as pd

file = 'AL010.xlsm'
file_1 = 'TT520b.xlsm'
file_4 = 'SG010.xlsm'

a = pd.read_excel(file, sheet_name='Data', usecols=['ARTNO', 'ARTNAME_UNICODE'])
b = pd.read_excel(file_1, sheet_name='Data', usecols=['ARTNO', 'TRANSTYPE', 'QTY', 'DATE'])
c = pd.read_excel(file_4, sheet_name='Data', usecols=['STORAGE_UNICODE', 'SGFLOCATION'])

a['ARTNO'] = a['ARTNO'].astype(str)
b['ARTNO'] = b['ARTNO'].astype(str)
c['STORAGE_UNICODE'] = c['STORAGE_UNICODE'].astype(str)

f = a.merge(b, how='left').merge(c, how='left', left_on=['ARTNO'], right_on=['STORAGE_UNICODE'])
f['ARTNO'] = f['ARTNO'].astype(int)
f = f[['ARTNO', 'ARTNAME_UNICODE', 'TRANSTYPE', 'QTY', 'DATE', 'SGFLOCATION']]
f.to_excel('how.xlsx', index="")

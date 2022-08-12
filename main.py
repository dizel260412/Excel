import pandas as pd
import pandas
# Assign spreadsheet filename to `file`
file = 'AL010.xlsm'
file_1 = 'TT520b.xlsm'
a = pd.read_excel(file, sheet_name='Data')
b = pd.read_excel(file_1, sheet_name='Data')


# print(a.shape)   ############ показывает, сколько в датафрейме строк и колонок
# print(a.columns)   ####### узнаем названия колонок
# print(a.dtypes)    ########### узнаем типы данных, находящихся в каждой колонке
# print(a.index)     #########
# print(a.describe())
# print(a['ARTNO'])          ########## Получаем данные определенного столбца
# print(a[['ARTNO','ARTNAME_UNICODE']])    ########## Получаем данные определенных столбцов [[...]]
# art = ['30313113']
# colum = ['ARTNO','ARTNAME_UNICODE']
# print(a.loc[colum] [art])


# art = a['ARTNO'] < 20000000 ######### Фильтор по колонке
# a = a.loc[art]
# print(art)

#
# a = a.query('ARTNO == 30313113')   ########## Поиск и вывод строки с артикулом, можноиспользовать несколько фильтров через &
# print(a)     #############Значение фильтра можно записать в переменную и вызывать через @ b = '30313113', ARTNO == @b


# b = pd.read_excel(file_1, sheet_name='Data')   ########### aнaлог SUMIF, по дате считаем qty
# b = b.groupby('DATE')['QTY'].sum()
# print(b)

# b = pd.read_excel(file_1, sheet_name='Data')   ########### aнaлог SUMIF, по дате считаем qty
# b = b.groupby(['DATE', 'TRANSTYPE'])['QTY'].sum()
# print(b)
#
# b = pd.read_excel(file_1, sheet_name='Data')   ########### aнaлог SUMIF, по дате считаем qty
# # b = b.query('TRANSTYPE == 260 & DATE == "2021-11-05"').groupby(['TRANSTYPE'])['QTY'].sum()  ##### Выводит сумму QTY по ТТ260 по дате 2021-11-05
# # print(b)
#
# art_260 = pd.merge(b, a, how='inner', left_index=True, right_on='ARTNO')
# print(art_260)

# b = pd.read_excel(file_1, sheet_name='Data')   ########### aнaлог SUMIFs, по дате считаем qty
# b = b.groupby(['DATE', 'TRANSTYPE'])['QTY'].agg(['sum', 'count']).sort_values(by='sum',ascending=False).head(10)   #### выводим sum и count QTY сортируем от большего к меньшему и выводим ТОП10
# print(b)





# # Load spreadsheet
# xl = pd.ExcelFile(file)
#
# # Print the sheet names
# #print(xl.sheet_names)
#
# # Load a sheet into a DataFrame by name: df1
# df1 = xl.parse('Data')
# print(df1)

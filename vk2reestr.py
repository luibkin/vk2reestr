# python 3.10.1

import pandas as pd
import os
import time

# Получение списка файлов в директории vk
files = os.listdir('vk')

# Объявление переменных
all_data = {}
date_data = {}
final_data = {}
r = 3
k = 0
p = 0

print('Формируется реестр из файлов:')
time.sleep(1)

for i in files:
    file = 'vk/' + i
    print(file)

# Считывание данных из файла в DataFrame df=vk
    df1 = pd.read_excel(file, sheet_name=0)

# Формирование словаря из ячеек с необходимыми данными.
# [5, 0], [7, 0], [9, 14], [27, 10] - координаты ячеек соответственно:
# номер акта, материал, дата, сертификаты. Можно добавить еще ячейки.
# Ячейка с сертификатом в данном случае в дальнейшем не используется.
    all_data[k] = {0: df1.iloc[5, 0], 1: df1.iloc[7, 0], 2: df1.iloc[9, 14], 3: df1.iloc[27, 10]}

# Формирование словаря из ячеек с датой (для упорядочивания результата
# по дате).
    date_data[k] = df1.iloc[9, 14]
    k = k + 1

# Получение списка ключей словаря из ячеек с датой при упорядочивании словаря
# по значениям даты.
for date in date_data:
    date_data[p] = date_data[p][6:10] + '.' + date_data[p][3:5] + '.' + date_data[p][:2]
    p = p + 1
sorted_date_data_keys = sorted(date_data, key=date_data.get)  # [1, 3, 2] - пример результата.

# Формирование словаря из ячеек с необходимыми данными под запись в реестр
# (с упорядочиванием по дате).
for key_b in sorted_date_data_keys:
    final_data[key_b] = {0: [('Акт результатов входного контроля - ' + str(all_data[key_b][1]) + ' от ' + str(all_data[key_b][2]))], 1: all_data[key_b][0][4:]}

# Запись в реестр.
with pd.ExcelWriter('./1.xlsx', engine='xlsxwriter') as writer:
    for key_c in sorted_date_data_keys:
        df2 = pd.DataFrame(final_data[key_c])
        df2.to_excel(writer, sheet_name="Sheet1", startrow=r, header=False, index=False)
        r = r + 1

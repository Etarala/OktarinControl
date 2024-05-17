import pandas as pd

# Загрузка данных из файла Excel
df = pd.read_excel('1.xlsx')

# Преобразование значений в числовой формат для столбцов 3 и 13
df.iloc[:, [2, 12]] = df.iloc[:, [2, 12]].apply(pd.to_numeric, errors='coerce')

# Переименование столбцов
df.rename(columns={
    'Unnamed: 2': '*Счет*',
    'Unnamed: 3': '*Владелец*',
    'Unnamed: 4': '*Карта*',
    'Клиент: Частные лица': '*АЗС*',
    'Unnamed: 7': '*Вид топлива*',
    'Unnamed: 10': '*Сумма покупки*',
    'Unnamed: 12': '*Начислено бонусов*',
    'Unnamed: 15': '*Тип операции*'
}, inplace=True)

# Фильтрация строк по условиям
filtered_df = df[df['*Тип операции*'] == 'Начисление'].copy()

# Удаление ненужных столбцов
columns_to_drop = [
    'Unnamed: 1',  # Столбец 1
    'Unnamed: 6',  # Столбец 6
    'Unnamed: 8',  # Столбец 8
    'Unnamed: 9',  # Столбец 9
    'Unnamed: 11', # Столбец 11
    'Unnamed: 13', # Столбец 13
    'Unnamed: 14', # Столбец 14
    'Unnamed: 16'  # Столбец 16
]
filtered_df.drop(columns=columns_to_drop, inplace=True)

# Подсчет количества повторений для каждого значения в столбце "*Счет*"
counts = filtered_df['*Счет*'].value_counts()

# Фильтрация по количеству повторений
filtered_df = filtered_df[filtered_df['*Счет*'].isin(counts.index[counts > 1])]

# Открытие текстового файла для записи результата
with open('Отчет.txt', 'w', encoding='utf-8') as file:
    # Группировка по столбцу "*Счет*" и итерация по группам
    for _, group_df in filtered_df.groupby('*Счет*'):
        # Запись строк с текущим повторяющимся значением в файл
        file.write(group_df.to_string(index=False))
        file.write('\n')
        # Запись суммы значений в столбце "*Начислено бонусов*" для текущей группы
        file.write(f"Всего начислено бонусов по счету: {group_df['*Начислено бонусов*'].sum()}\n\n")


# Вывод в консоль
# Группировка по столбцу "Счет" и итерация по группам
#for _, group_df in filtered_df.groupby('*Счет*'):
    # Вывод строк с текущим повторяющимся значением
#    print(group_df.to_string(index=False))
    # Вывод суммы значений в столбце "Количество бонусов" для текущей группы
#    print(f"Всего начислено бонусов по счету: {group_df['*Начислено бонусов*'].sum()}\n")
#input()


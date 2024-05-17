import pandas as pd

# Загрузка данных из файла Excel
df = pd.read_excel('1.xlsx')

# Преобразование значений в числовой формат для столбцов 3 и 13
df.iloc[:, [2, 12]] = df.iloc[:, [2, 12]].apply(pd.to_numeric, errors='coerce')

# Фильтрация строк по условиям
filtered_df = df[df.iloc[:, 15] == 'Начисление']

# Удаление столбцов с именами
columns_to_drop = [
    df.columns[1],  # Столбец 1
    df.columns[6],  # Столбец 6
    df.columns[8],  # Столбец 8
    df.columns[9],  # Столбец 9
    df.columns[11], # Столбец 11
    df.columns[13], # Столбец 13
    df.columns[14], # Столбец 14
    df.columns[16]  # Столбец 16
]
filtered_df = filtered_df.drop(columns=columns_to_drop)

# Группировка по третьему столбцу (его имя) и итерация по группам
group_column_name = df.columns[2]
sum_column_name = df.columns[12]

for _, group_df in filtered_df.groupby(filtered_df[group_column_name]):
    # Вывод строк с текущим повторяющимся значением
    print(group_df.to_string(index=False))
    # Вывод суммы значений в столбце с нужным именем для текущей группы
    print(f"Сумма значений в столбце 13: {group_df[sum_column_name].sum()}\n")

import pandas as pd

# Загрузка данных из файла Excel
df = pd.read_excel('1.xlsx')

# Преобразование значений в числовой формат для столбцов 3 и 13
df.iloc[:, [2, 12]] = df.iloc[:, [2, 12]].apply(pd.to_numeric, errors='coerce')

# Фильтрация строк по условиям
filtered_df = df[df.iloc[:, 15] == 'Начисление']

# Группировка по третьему столбцу и итерация по группам
for _, group_df in filtered_df.groupby(filtered_df.iloc[:, 2]):
    # Вывод строк с текущим повторяющимся значением
    print(group_df.to_string(index=False))
    # Вывод суммы значений в тринадцатом столбце для текущей группы
    print(f"Сумма значений в столбце 13: {group_df.iloc[:, 12].sum()}\n")

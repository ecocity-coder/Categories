![alt text](https://github.com/ecocity-coder/Categories/blob/main/2024-05-07_10-33-51.png)

import pandas as pd
file_path = "C:/Users/USER/Downloads/апрель.xlsx"
xlsx_file = pd.ExcelFile(file_path)
# Цикл по всем листам в файле
for sheet_name in xlsx_file.sheet_names:
    # Чтение данных из листа
    df = pd.read_excel(xlsx_file, sheet_name)
    
    # Запись данных в новый файл
    output_file = f'{sheet_name}.xlsx'
    df.to_excel(output_file, index=False)


<h2><span style="color: #0000ff;"><strong>для удобства работы я разделил таблицу на листы и листы сохранил как отдельные файлы</strong></span></h2>

<h2><span style="color: #0000ff;"><strong>1. Установлю, за счет каких товарных категорий получена наибольшая валовая прибыль.</strong></span></h2>

file_path = 'C:/Users/USER/Факт продаж.xlsx'
df = pd.read_excel(file_path)

# Группировка данных по товарной категории и подсчет суммарной валовой прибыли
grouped_data = df.groupby('Товарная категория')['Валовая прибыль'].sum()

# Вывод результатов
print(grouped_data)

# Создание нового датафрейма из сгруппированных данных
result_df = pd.DataFrame(grouped_data).reset_index()
result_df.columns = ['Товарная категория', 'Сумма валовой прибыли']

# Сохранение результатов в новую таблицу xlsx
result_file_path = 'C:/Users/USER/gross_profit.xlsx'
result_df.to_excel(result_file_path, index=False)

print("Данные сохранены в новую таблицу xlsx.")

# Сортировка по убыванию валовой прибыли и выбор первых пяти категорий
top_categories = grouped_data.sort_values(ascending=False).head(5)

# Вывод результатов
print(top_categories)
top_categories_df = pd.DataFrame(top_categories).reset_index()
top_categories_df.columns = ['Товарная категория', 'Сумма валовой прибыли']
top_categories_file_path = 'C:/Users/USER/top_gross_profit_categories.xlsx'
top_categories_df.to_excel(top_categories_file_path, index=False)
print("Новая таблица xlsx с первыми пятью категориями по валовой прибыли сохранена.")

<h2><span style="color: #0000ff;"><strong>Получено 5 товарных категорий, принесших наибольшую валовую прибыль.</strong></span></h2>

import matplotlib.pyplot as plt
# Проверка на отрицательные значения и замена их на 0
grouped_data[grouped_data < 0] = 0
# Создание круговой диаграммы
plt.figure(figsize=(10, 7))
plt.pie(grouped_data, labels=grouped_data.index, autopct='%1.1f%%')
plt.axis('equal')
plt.title('Распределение валовой прибыли по товарным категориям')
plt.show()

<h2><span style="color: #0000ff;"><strong>Для удобства визуализации выберу первые пять товарных категорий, принесших наибольшую валовую прибыль.</strong></span></h2>

# Выбор только топ-5 категорий
top_categories = grouped_data.sort_values(ascending=False).head(5)
# Проверка на отрицательные значения и замена их на 0
grouped_data[grouped_data < 0] = 0

# Создание круговой диаграммы
plt.figure(figsize=(10, 7))
plt.pie(top_categories, labels=top_categories.index, autopct='%1.1f%%')
plt.axis('equal')
plt.title('Распределение валовой прибыли по первым пять товарным категориям')
plt.show()

<h2><span style="color: #0000ff;"><strong>2. Установлю подразделение, которое продало более прочих товарных категорий с наибольшей валовой прибылью.</strong></span></h2>

# Группировка данных по двум колонкам и суммирование валовой прибыли
grouped = df.groupby(['Заказ клиента / Реализация.Соглашение.Менеджер.Подразделение', 'Товарная категория'])['Валовая прибыль'].sum().reset_index()

# Отбор товарных категорий с максимальной валовой прибылью
max_profit_categories = grouped.groupby('Товарная категория')['Валовая прибыль'].transform(max) == grouped['Валовая прибыль']

# Отбор подразделений, которые продали больше всего товарных категорий с максимальной валовой прибылью
result = grouped[max_profit_categories].groupby('Заказ клиента / Реализация.Соглашение.Менеджер.Подразделение')['Товарная категория'].count().idxmax()

print(f"Подразделение, которое продало больше всего товарных категорий с максимальной валовой прибылью: {result}")


<h2><span style="color: #0000ff;"><strong>Покажем это на графике.</strong></span></h2>

sales_data = grouped[max_profit_categories].groupby('Заказ клиента / Реализация.Соглашение.Менеджер.Подразделение')['Товарная категория'].count().reset_index()
sales_data.columns = ['Заказ клиента / Реализация.Соглашение.Менеджер.Подразделение', 'количество_товарных_категорий']
# Создание файла Excel
sales_data.to_excel('Продажи_товарных_категорий_по_подразделениям.xlsx', index=False)

print("Файл 'Продажи_товарных_категорий_по_подразделениям.xlsx' успешно создан.")

# Построение графика
plt.figure(figsize=(10, 6))
plt.bar(sales_data['Заказ клиента / Реализация.Соглашение.Менеджер.Подразделение'], sales_data['количество_товарных_категорий'], color='skyblue')
plt.xlabel('Подразделение')
plt.ylabel('Количество товарных категорий с макс. валовой прибылью')
plt.title('Продажи товарных категорий с макс. валовой прибылью по подразделениям')
plt.xticks(rotation=45)
plt.show()


<h2><span style="color: #0000ff;"><strong>3. установим, какие планы были у подразделений по валовой прибыли.</strong></span></h2>

# установим, какие планы были у подразделений по валовой прибыли
file_path = 'C:/Users/USER/План продаж.xlsx'
df1 = pd.read_excel(file_path)
# для удобства восприятия информации приведем числовой формат колонки к одному знаку после запятой
df1['Валовая прибыль, руб.'] = df1['Валовая прибыль, руб.'].round(1)
result = df1.groupby('Подразделение')['Валовая прибыль, руб.'].sum()
result.to_excel('dept_plan.xlsx')

print(result)

# посмотрим, какую фактически валовую прибыль принесли подразделения
file_path = 'C:/Users/USER/Факт продаж.xlsx'
df2 = pd.read_excel(file_path)
result = df2.groupby('Заказ клиента / Реализация.Соглашение.Менеджер.Подразделение')['Валовая прибыль'].sum()

print(result)

# Чтение данных из файлов "План продаж.xlsx" и "Факт продаж.xlsx"
file_path_plan = 'C:/Users/USER/План продаж.xlsx'
file_path_fact = 'C:/Users/USER/Факт продаж.xlsx'

df1 = pd.read_excel(file_path_plan)
df2 = pd.read_excel(file_path_fact)

# Группировка данных и вычисление суммарной валовой прибыли
plan_profit = df1.groupby('Подразделение')['Валовая прибыль, руб.'].sum().reset_index()
fact_profit = df2.groupby('Заказ клиента / Реализация.Соглашение.Менеджер.Подразделение')['Валовая прибыль'].sum().reset_index()

# Объединение данных о плановой и фактической прибыли по подразделениям
result = pd.merge(plan_profit, fact_profit, left_on='Подразделение', right_on='Заказ клиента / Реализация.Соглашение.Менеджер.Подразделение', how='outer')
result.columns = ['Подразделение', 'Плановая валовая прибыль, руб.', 'Фактическая валовая прибыль']

# Создание нового Excel файла "dept_profit.xlsx" с результатами
output_file_path = 'C:/Users/USER/dept_profit.xlsx'
result.to_excel(output_file_path, index=False)

print("Файл dept_profit.xlsx успешно создан с информацией о фактической валовой прибыли по подразделениям.")

<h2><span style="color: #0000ff;"><strong>из полученных данных видно что "Отдел продаж Хорека" имел план продаж почти втрое больше фактически исполненного. "Отдел продаж г Москва" фактически исполнил продаж в 2,5 раза менее запланированного. Отдел территориальых продаж имел план в 14 раз больше фактически исполненного.</strong></span></h2>

<h2><span style="color: #0000ff;"><strong>5. рассчитаем средний чек по менеджерам за квартал.</strong></span></h2>

file_path = 'C:/Users/USER/Продажи 1 кв 2023.xlsx'
df3 = pd.read_excel(file_path)
# Проверка наличия заголовка "Выручка" в данных и удаление
if 'Выручка' in df3['Выручка'].values:
    df3 = df3[df3['Выручка'] != 'Выручка']
# Замена пропущенных значений в столбце "Выручка" на 0 
df3['Выручка'] = df3['Выручка'].fillna(0)
# Преобразование столбца "Выручка" в числовой тип данных
df3['Выручка'] = pd.to_numeric(df3['Выручка'])
avg_revenue = df3.groupby('Менеджер')['Выручка'].mean()
# Округление средней выручки до трех знаков
avg_revenue = avg_revenue.round(3)

print(avg_revenue)




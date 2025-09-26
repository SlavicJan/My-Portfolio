import pandas as pd
import numpy as np
from sqlalchemy import create_engine
import warnings
# warnings.filterwarnings('ignore')

# Настройки подключения к БД (замените на свои)
# DB_CONFIG = {
#     'host': 'your_host',
#     'database': 'your_db',
#     'user': 'your_user',
#     'password': 'your_password'
# }

# Подключение к базе данных
# engine = create_engine(f'postgresql://{DB_CONFIG["user"]}:{DB_CONFIG["password"]}@{DB_CONFIG["host"]}/{DB_CONFIG["database"]}')

# Загрузка данных
# query = """
# SELECT
#     dr_ndrugs,
#     dr_dat,
#     dr_kol,
#     dr_croz,
#     dr_czak,
#     dr_sdisc
# FROM sales
# """

# df = pd.read_sql(query, engine) ////

df = pd.read_csv(r"C:\Users\Slavic Asus\OneDrive\Документы\GitHub\My-Projects\sales_202509091515.csv", engine='python')

# # Предварительная обработка данных

print(df.isnull().sum())
print(df.duplicated().sum())
print(df.shape)
print(df.describe())
IQR = df.quantile([0.25, 0.75])
IQR = IQR[1] - IQR[0]
df = df[~((df < (IQR * 1.5)) | (df > (IQR * 1.5))).any(axis=1)]
# print(df.shape)


# df['revenue'] = df['dr_kol'] * df['dr_croz'] - df['dr_sdisc']
# df['profit'] = df['dr_kol'] * (df['dr_croz'] - df['dr_czak']) - df['dr_sdisc']
#
# # ABC анализ
# abc_data = df.groupby('dr_ndrugs').agg({
#     'dr_kol': 'sum',
#     'revenue': 'sum',
#     'profit': 'sum'
# }).reset_index()
#
# # Функция для расчета ABC категорий
# def calculate_abc(series):
#     sorted_series = series.sort_values(ascending=False)
#     cumulative_percentage = sorted_series.cumsum() / sorted_series.sum()
#     abc = pd.cut(cumulative_percentage,
#                  bins=[0, 0.8, 0.95, 1],
#                  labels=['A', 'B', 'C'],
#                  include_lowest=True)
#     return abc
#
# abc_data['abc_amount'] = calculate_abc(abc_data['dr_kol'])
# abc_data['abc_revenue'] = calculate_abc(abc_data['revenue'])
# abc_data['abc_profit'] = calculate_abc(abc_data['profit'])
#
# # XYZ анализ
# df['dr_dat'] = pd.to_datetime(df['dr_dat'])
# df['yw'] = df['dr_dat'].dt.strftime('%Y-%W')
# xyz_data = df.groupby(['dr_ndrugs', 'yw'])['dr_kol'].sum().reset_index()
#
# # Расчет коэффициента вариации
# xyz_calc = xyz_data.groupby('dr_ndrugs').agg(
#     weekly_sales=('dr_kol', 'std'),
#     avg_sales=('dr_kol', 'mean'),
#     n_weeks=('yw', 'nunique')
# ).reset_index()
#
# xyz_calc['cv'] = xyz_calc['weekly_sales'] / xyz_calc['avg_sales']
#
# # Определение XYZ категорий
# conditions = [
#     (xyz_calc['cv'] <= 0.1) & (xyz_calc['n_weeks'] >= 4) & (xyz_calc['n_weeks'] < 7),
#     (xyz_calc['cv'] > 0.1) & (xyz_calc['cv'] <= 0.25) & (xyz_calc['n_weeks'] >= 4) & (xyz_calc['n_weeks'] < 7),
#     (xyz_calc['cv'] > 0.25) & (xyz_calc['n_weeks'] >= 4) & (xyz_calc['n_weeks'] < 7)
# ]
#
# choices = ['X', 'Y', 'Z']
# xyz_calc['xyz'] = np.select(conditions, choices, default=None)
#
# # Объединение результатов
# final_result = abc_data.merge(
#     xyz_calc[['dr_ndrugs', 'xyz']],
#     on='dr_ndrugs',
#     how='left'
# )
#
# # Сохранение результатов
# final_result.to_csv('abc_xyz_analysis.csv', index=False)
# print("Анализ завершен. Результаты сохранены в abc_xyz_analysis.csv")
#
# # Дополнительная статистика
# print("\nСтатистика анализа:")
# print(f"Всего товаров: {len(final_result)}")
# print(f"Товаров с XYZ классификацией: {final_result['xyz'].notnull().sum()}")
# print("\nРаспределение по категориям:")
# print("По количеству:")
# print(final_result['abc_amount'].value_counts())
# print("\nПо прибыли:")
# print(final_result['abc_profit'].value_counts())
# print("\nПо выручке:")
# print(final_result['abc_revenue'].value_counts())
# print("\nПо стабильности продаж (XYZ):")
# print(final_result['xyz'].value_counts())
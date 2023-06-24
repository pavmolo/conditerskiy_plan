import streamlit as st
import numpy as np
import pandas as pd

st.markdown('Загрузите файл с мастер данными')
master_data_file = st.file_uploader("Выберите XLSX файл с мастер данными", accept_multiple_files=False)
st.markdown('Загрузите файл с планом')
plan_file = st.file_uploader("Выберите XLSX файл с планом", accept_multiple_files=False)

cycle_time_table = pd.read_excel(master_data_file, sheet_name='cycle_time_table')
time_mode = pd.read_excel(master_data_file, sheet_name='time_mode')
cream_data = pd.read_excel(master_data_file, sheet_name='cream_data')
current_plan = pd.read_excel(plan_file, sheet_name='current_date')

# Объединяем таблицы с циклами и планом
operation_plan_for_all_calls = current_plan.merge(cycle_time_table, on=['sku', 'operation'], how='left')
operation_plan_for_all_calls['production_duration_sec'] = operation_plan_for_all_calls['quantity'] * operation_plan_for_all_calls['cycle_time_sec']

# Дополняем таблицу временного режима столбцами временными границами с начала смены
time_mode['duration_cumulative'] = time_mode['duration'].cumsum()
time_mode['duration_cumulative_start'] = pd.Series([0])._append(time_mode['duration_cumulative'].shift(1).iloc[1:]).astype('int')

# Перечисляем все ячейки из плана
cell_list = operation_plan_for_all_calls['cell'].unique()

# Создаем обособленные по ячейкам планы
cell_plan_list = []
for i in cell_list:
  cell_plan_list.append(operation_plan_for_all_calls[operation_plan_for_all_calls['cell'] == i])
for i in cell_plan_list:
  i['cycle_time_sec_cumulative'] = i['production_duration_sec'].cumsum()
  i['duration_cumulative_start'] = pd.Series([0])._append(i['cycle_time_sec_cumulative'].shift(1).iloc[1:]).astype('int')

# Объединяем таблицы с циклами и планом
operation_plan_for_all_calls = current_plan.merge(cycle_time_table, on=['sku', 'operation'], how='left')
operation_plan_for_all_calls['production_duration_sec'] = operation_plan_for_all_calls['quantity'] * operation_plan_for_all_calls['cycle_time_sec']

# Дополняем таблицу временного режима столбцами временными границами с начала смены
time_mode['duration_cumulative'] = time_mode['duration'].cumsum()
time_mode['duration_cumulative_start'] = pd.Series([0])._append(time_mode['duration_cumulative'].shift(1).iloc[1:]).astype('int')

# Перечисляем все ячейки из плана
cell_list = operation_plan_for_all_calls['cell'].unique()

# Создаем обособленные по ячейкам планы
cell_plan_list = []
for i in cell_list:
  cell_plan_list.append(operation_plan_for_all_calls[operation_plan_for_all_calls['cell'] == i])
for i in cell_plan_list:
  i['cycle_time_sec_cumulative'] = i['production_duration_sec'].cumsum()
  i['duration_cumulative_start'] = pd.Series([0])._append(i['cycle_time_sec_cumulative'].shift(1).iloc[1:]).astype('int')

# Производим основные расчеты через создание посекундки
df_list = []
for i in cell_plan_list:
  df = pd.DataFrame({'Номер секунды': range(1, time_mode['duration_cumulative'].iloc[-1] + 1)})
  def find_time_window(row):
    for j in range(len(time_mode)):
        if time_mode['duration_cumulative_start'][j] <= row['Номер секунды'] <= time_mode['duration_cumulative'][j]:
            return time_mode['start'][j]
    return None
  def find_sku(row):
    for j in i.index:
        if i['duration_cumulative_start'][j] <= row['Номер секунды'] <= i['cycle_time_sec_cumulative'][j]:
            return i['sku'][j]
    return None
  def find_operation(row):
    for j in i.index:
        if i['duration_cumulative_start'][j] <= row['Номер секунды'] <= i['cycle_time_sec_cumulative'][j]:
            return i['operation'][j]
    return None

  df['time_window'] = df.apply(find_time_window, axis=1)
  df['sku'] = df.apply(find_sku, axis=1)
  df['operation'] = df.apply(find_operation, axis=1)
  pivot_table = pd.pivot_table(df, index=['time_window', 'sku', 'operation'], aggfunc='count').reset_index().merge(operation_plan_for_all_calls[['cell', 'sku', 'operation', 'cycle_time_sec']], on=['sku', 'operation'], how='left')
  pivot_table['quantity'] = (pivot_table['Номер секунды'] / pivot_table['cycle_time_sec']).astype('int')
  df_list._append(pivot_table[['cell', 'time_window', 'sku', 'operation', 'quantity']])
  df_list_with_cream = []
  for i in df_list:
    cake_plan_with_cream = df_list[0].merge(cream_data, on=['sku', 'operation'], how='left')
    cake_plan_with_cream['cream_plan'] = cake_plan_with_cream['quantity'] * cake_plan_with_cream['gr']
    df_list_with_cream._append(cake_plan_with_cream)
  df_cream_list = []
  for i in df_list_with_cream:
    df_by_creams_and_time = pd.pivot_table(i, index=['time_window', 'raw_materials'], values='cream_plan', aggfunc='sum').reset_index()
    df_cream_list._append(df_by_creams_and_time)
  merged_df = pd.concat(df_cream_list, axis=0)
  cream_time = pd.pivot_table(merged_df, index=['time_window', 'raw_materials'], values='cream_plan', aggfunc='sum').reset_index()
  

if master_data_file:
  if plan_file:
    for i in df_list:
      st.dataframe(i)
    st.dataframe(cream_time)

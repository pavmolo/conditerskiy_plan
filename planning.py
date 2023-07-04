import streamlit as st
import numpy as np
import pandas as pd
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb

st.markdown('''<a href="http://kaizen-consult.ru/"><img src='https://www.kaizen.com/images/kaizen_logo.png' style="width: 50%; margin-left: 25%; margin-right: 25%; text-align: center;"></a><p>''', unsafe_allow_html=True)
st.markdown('''<h1>Приложение для разбивки плана по ячейкам и определения потребности в сырье по часам</h1>''', unsafe_allow_html=True)
col1, col2 = st.columns(2)

with col1:
  st.markdown('''<h3>Загрузите файл с мастер данными</h3>''', unsafe_allow_html=True)
  master_data_file = st.file_uploader("Выберите XLSX файл с мастер данными", accept_multiple_files=False)
with col2:
  st.markdown('''<h3>Загрузите файл с планом</h3>''', unsafe_allow_html=True)
  plan_file = st.file_uploader("Выберите XLSX файл с планом", accept_multiple_files=False)
if master_data_file:
  if plan_file:
    cycle_time_table = pd.read_excel(master_data_file, sheet_name='cycle_time_table')
    time_mode = pd.read_excel(master_data_file, sheet_name='time_mode')
    cream_data = pd.read_excel(master_data_file, sheet_name='cream_data')
    current_plan = pd.read_excel(plan_file, sheet_name='current_date')
    cat_values = time_mode['start']
    time_mode['start'] = pd.Categorical(time_mode['start'], categories=cat_values, ordered=True)
    
    #______________________________________________________________________
    
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
      pivot_table['time_window'] = pd.Categorical(pivot_table['time_window'], categories=cat_values, ordered=True)
      pivot_table = pivot_table.sort_values('time_window')
      df_list.append(pivot_table[['cell', 'time_window', 'sku', 'operation', 'quantity']])
    
    #____________________________________________________
      
      df_list_with_cream = []
      for i in df_list:
        cake_plan_with_cream = i.merge(cream_data, on=['sku', 'operation'], how='left')
        cake_plan_with_cream['cream_plan'] = cake_plan_with_cream['quantity'] * cake_plan_with_cream['gr']
        df_list_with_cream.append(cake_plan_with_cream)
      df_cream_list = []
      for i in df_list_with_cream:
        df_by_creams_and_time = pd.pivot_table(i, index=['time_window', 'raw_materials'], values='cream_plan', aggfunc='sum').reset_index()
        df_cream_list.append(df_by_creams_and_time)
      merged_df = pd.concat(df_cream_list, axis=0)
      cream_time = pd.pivot_table(merged_df, index=['time_window', 'raw_materials'], values='cream_plan', aggfunc='sum').reset_index()
      cream_time['time_window'] = pd.Categorical(cream_time['time_window'], categories=cat_values, ordered=True)
      cream_time = cream_time.sort_values('time_window')
      cream_time = cream_time[cream_time['cream_plan'] != 0]
    

    st.title('План по ячейкам')
    for i in df_list:
      st.dataframe(i)
    st.title('Потребность в сырье')
    st.dataframe(cream_time)
    def to_excel():
      output = BytesIO()
      writer = pd.ExcelWriter(output, engine='xlsxwriter')
      cream_time.to_excel(writer, index=False, sheet_name='cream_time')
      with writer as w:
        for i in df_list:
          i.to_excel(w, sheet_name=i['cell'][0].replace('/', '-'))
      writer._save()
      processed_data = output.getvalue()
      return processed_data
    df_xlsx = to_excel()
    st.download_button(label='📥 Скачать план в Excel',
                       data=df_xlsx ,
                       file_name= 'Safia_Plan.xlsx')

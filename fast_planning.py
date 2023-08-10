import streamlit as st
import numpy as np
import pandas as pd
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb


def distribute_operations(time_mode, cycles, plan):
    # Объединяем данные из plan и cycles по операциям
    merged_plan = plan.merge(cycles, on='operation', how='left')

    # Вычисляем общее время выполнения для каждой операции
    merged_plan['total_time'] = merged_plan['cycle_time'] * merged_plan['quantity']

    # Получаем уникальные ячейки
    unique_cells = merged_plan['cell'].unique()

    dfs = []

    for cell in unique_cells:
        cell_operations = merged_plan[merged_plan['cell'] == cell].copy()
        time_mode_copy = time_mode.copy()
        time_mode_copy['remaining_time'] = time_mode_copy['working_seconds']

        cell_result = []

        for _, time_row in time_mode_copy.iterrows():
            for _, operation_row in cell_operations.iterrows():
                operation = operation_row['operation']
                total_time = operation_row['total_time']
                cycle_time = operation_row['cycle_time']

                if total_time <= 0:
                    continue

                # Вычисляем, сколько операций можно выполнить в текущем часовом интервале
                operations_count = min(total_time // cycle_time, time_row['remaining_time'] // cycle_time)

                if operations_count > 0:
                    cell_result.append({
                        'hour_interval': time_row['hour_interval'],
                        'operation': operation,
                        'operations_count': operations_count
                    })

                    allocated_time = operations_count * cycle_time
                    total_time -= allocated_time
                    time_row['remaining_time'] -= allocated_time
                    operation_row['total_time'] = total_time
                    # Обновляем количество операций в plan
                    idx = cell_operations[cell_operations['operation'] == operation].index[0]
                    cell_operations.at[idx, 'total_time'] = total_time
        if cell_result:  # Проверяем, не пуст ли список
            dfs.append(pd.DataFrame(cell_result).sort_values(by='hour_interval'))


    return dfs










st.markdown('''<a href="http://kaizen-consult.ru/"><img src='https://www.kaizen.com/images/kaizen_logo.png' style="width: 50%; margin-left: 25%; margin-right: 25%; text-align: center;"></a><p>''', unsafe_allow_html=True)
st.markdown('''<h1>Приложение для разбивки плана по ячейкам и определения потребности в сырье по часам</h1>''', unsafe_allow_html=True)
col1, col2 = st.columns(2)







with col1:
  st.markdown('''<h3>Файл с мастер данными</h3>''', unsafe_allow_html=True)
  master_data_file = st.file_uploader("Выберите XLSX файл с мастер данными", accept_multiple_files=False)
with col2:
  st.markdown('''<h3>Файл с планом</h3>''', unsafe_allow_html=True)
  plan_file = st.file_uploader("Выберите XLSX файл с планом", accept_multiple_files=False)
if master_data_file:
  if plan_file:
    cycle_time_table = pd.read_excel(master_data_file, sheet_name='cycle_time_table')
    cycle_time_table['cycle_time_sec'] = cycle_time_table['cycle_time_sec'].astype('int')
    time_mode = pd.read_excel(master_data_file, sheet_name='time_mode')
    cream_data = pd.read_excel(master_data_file, sheet_name='cream_data')
    current_plan = pd.read_excel(plan_file, sheet_name='current_date')
    cat_values = time_mode['start']
    cat_values_plan = current_plan['sku']
    time_mode['start'] = pd.Categorical(time_mode['start'], categories=cat_values, ordered=True)
    current_plan['sku'] = pd.Categorical(current_plan['sku'], categories=cat_values_plan, ordered=True)


    # Пример использования
    time_mode_data = {
        'hour_interval': time_mode['start'],
        'working_seconds': time_mode['duration']
    }
    
    cycles_data = {
        'operation': cycle_time_table['sku'],
        'cycle_time': cycle_time_table['cycle_time_sec'],
        'cell': cycle_time_table['cell'],
    }
    
    plan_data = {
        'operation': current_plan['sku'],
        'quantity': current_plan['quantity']
    }
    
    time_mode_df = pd.DataFrame(time_mode_data)
    cycles_df = pd.DataFrame(cycles_data)
    plan_df = pd.DataFrame(plan_data)
    #______________________________________________________________________
    


    dataframes = distribute_operations(time_mode_df, cycles_df, plan_df)
    dfdf = []
    for df in dataframes:
      dfdf.append(df.sort_values(by=['operation', 'hour_interval']))
    with st.expander("Посмотреть таблицы"):
      st.title('План по ячейкам')
      for i in dfdf:
        st.dataframe(i)
    def to_excel():
      output = BytesIO()
      writer = pd.ExcelWriter(output, engine='xlsxwriter')
      with writer as w:
        for i in dfdf:
          i.to_excel(w, sheet_name=i['cell'][0].replace('/', '-'))
      writer._save()
      processed_data = output.getvalue()
      return processed_data
    df_xlsx = to_excel()
    st.download_button(label='📥 Скачать план в Excel',
                       data=df_xlsx ,
                       file_name= 'Safia_Plan.xlsx')

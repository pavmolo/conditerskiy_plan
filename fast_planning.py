import streamlit as st
import numpy as np
import pandas as pd
from io import BytesIO

def distribute_operations(time_mode_var, cycles, plan):
    merged_plan = plan.merge(cycles, on='operation', how='left')
    merged_plan['total_time'] = merged_plan['cycle_time'] * merged_plan['quantity']
    unique_cells = merged_plan['cell'].unique()
    dfs = []

    for cell in unique_cells:
        cell_operations = merged_plan[merged_plan['cell'] == cell].copy()
        time_mode_copy = time_mode_var.copy()
        time_mode_copy['remaining_time'] = time_mode_copy['working_seconds']
    
        cell_result = []
    
        time_index = 0
        for _, operation_row in cell_operations.iterrows():
            operation = operation_row['operation']
            total_time = operation_row['total_time']
            cycle_time = operation_row['cycle_time']
    
            if total_time <= 0:
                continue
    
            while total_time > 0 and time_index < len(time_mode_copy):
                time_row = time_mode_copy.iloc[time_index]
                
                if time_row['remaining_time'] < cycle_time:
                    time_index += 1
                    if time_index >= len(time_mode_copy):
                        break
                    time_row = time_mode_copy.iloc[time_index]
                    continue
    
                operations_count = np.floor(min(total_time / cycle_time, time_row['remaining_time'] / cycle_time))
    
                if operations_count > 0:
                    cell_result.append({
                        'hour_interval': time_row['hour_interval'],
                        'operation': operation,
                        'operations_count': operations_count
                    })
    
                    allocated_time = operations_count * cycle_time
                    total_time -= allocated_time
                    time_mode_copy.at[time_index, 'remaining_time'] -= allocated_time
    
                if total_time <= 0:
                    break
    
        if cell_result:
            df = pd.DataFrame(cell_result)
            df['operation'] = pd.Categorical(df['operation'], categories=current_plan['sku'], ordered=True)
            df['hour_interval'] = pd.Categorical(df['hour_interval'], categories=time_mode['start'], ordered=True)
            df['cell'] = cell
            dfs.append(df.sort_values(by=['operation', 'hour_interval']))

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

if master_data_file and plan_file:
    cycle_time_table = pd.read_excel(master_data_file, sheet_name='cycle_time_table')
    cycle_time_table['cycle_time_sec'] = cycle_time_table['cycle_time_sec'].astype('int')
    time_mode = pd.read_excel(master_data_file, sheet_name='time_mode')
    current_plan = pd.read_excel(plan_file, sheet_name='current_date')

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

    dataframes = distribute_operations(time_mode_df, cycles_df, plan_df)

    with st.expander("Посмотреть таблицы"):
        st.title('План по ячейкам')
        for df in dataframes:
            cell_name = df['cell'].iloc[0]
            st.markdown(f"### {cell_name}")
            st.dataframe(df.drop(columns=['cell']))

    try:
        cream_data = pd.read_excel(master_data_file, sheet_name='cream_data')
        st.write("Данные о сырье до объединения:")
        st.dataframe(cream_data)
    except Exception as e:
        st.warning("Не удалось загрузить данные о сырье. Убедитесь, что в файле есть лист 'cream_data'.")
        cream_data = pd.DataFrame(columns=['sku', 'operation', 'raw_materials', 'gr'])

    # Проверим содержимое all_data
    st.write("Содержимое all_data:")
    st.dataframe(pd.concat(dataframes))

    # Объединяем данные без использования категориальных данных
    all_data_non_cat = pd.concat(dataframes).astype(str)
    # Объединяем по столбцам 'operation' и 'sku'
    merged_data = all_data_non_cat.merge(cream_data, left_on='operation', right_on='sku', how='inner')
    merged_data['total_gr'] = merged_data['operations_count'].astype(float) * merged_data['gr'].astype(float)
    raw_materials_df = merged_data.groupby(['hour_interval', 'raw_materials'])['total_gr'].sum().reset_index()
    raw_materials_df['hour_interval'] = pd.Categorical(df['hour_interval'], categories=time_mode['start'], ordered=True)
    raw_materials_df = raw_materials_df.sort_values(by=['hour_interval'])
    st.write("Данные о сырье после объединения:")
    st.dataframe(raw_materials_df)

    def to_excel():
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        with writer as w:
            for df in dataframes:
                df.to_excel(w, sheet_name=df['cell'].iloc[0].replace('/', '-'))
        writer._save()
        return output.getvalue()

    df_xlsx = to_excel()
    st.download_button(label='📥 Скачать план в Excel', data=df_xlsx, file_name='Safia_Plan.xlsx')

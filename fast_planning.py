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
    
                if operations_count == 0:
                    time_index += 1
                    continue
    
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
def get_final_times(dataframes):
    final_times_list = []
    for df in dataframes:
        cell_name = df['cell'].iloc[0]
        final_time_window = df['hour_interval'].iloc[-1]
        final_times_list.append({
            'cell': cell_name,
            'final_time_window': final_time_window
        })
    return pd.DataFrame(final_times_list)

col1, col2 = st.columns(2)

with col1:
    st.markdown('''<h3>Файл с мастер данными</h3>''', unsafe_allow_html=True)
    master_data_file = st.file_uploader("Выберите XLSX файл с мастер данными", accept_multiple_files=False)

with col2:
    st.markdown('''<h3>Файл с планом</h3>''', unsafe_allow_html=True)
    plan_file = st.file_uploader("Выберите XLSX файл с планом", accept_multiple_files=False)

if master_data_file and plan_file:
    #st.write("Файлы с мастер данными и планом успешно загружены.")
    cycle_time_table = pd.read_excel(master_data_file, sheet_name='cycle_time_table')
    time_mode = pd.read_excel(master_data_file, sheet_name='time_mode')
    current_plan = pd.read_excel(plan_file, sheet_name='current_date')
    #st.write("Данные из файлов успешно прочитаны.")
    
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
    plan_df['quantity'] = plan_df['quantity'].round().astype(int)

    #st.write("Данные успешно преобразованы в датафреймы.")
    
    dataframes = distribute_operations(time_mode_df, cycles_df, plan_df)

    with st.expander("Посмотреть почасовые планы по ячейкам"):
        st.title('План по ячейкам')
        for df in dataframes:
            cell_name = df['cell'].iloc[0]
            st.markdown(f"### {cell_name}")
            st.dataframe(df.drop(columns=['cell']))
            
    with st.expander("Посмотреть данные по сырью и время окончания работы по ячейкам"):
        st.title('План по сырью')
        try:
            cream_data = pd.read_excel(master_data_file, sheet_name='cream_data')
        except:
            cream_data = pd.DataFrame(columns=['sku', 'operation', 'raw_materials', 'gr'])
    
        all_data_non_cat = pd.concat(dataframes).astype(str)
        merged_data = all_data_non_cat.merge(cream_data, left_on='operation', right_on='sku', how='inner')
        merged_data['total_gr'] = merged_data['operations_count'].astype(float) * merged_data['gr'].astype(float)
        raw_materials_df = merged_data.groupby(['hour_interval', 'raw_materials'])['total_gr'].sum().reset_index()
        raw_materials_df['hour_interval'] = pd.Categorical(raw_materials_df['hour_interval'], categories=time_mode['start'], ordered=True)
        raw_materials_df = raw_materials_df.sort_values(by=['hour_interval', 'raw_materials'])
        st.dataframe(raw_materials_df)
        
        final_times = get_final_times(dataframes)
        st.dataframe(final_times)

    def to_excel():
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        with writer as w:
            for df in dataframes:
                df.to_excel(w, sheet_name=df['cell'].iloc[0].replace('/', '-'))
            raw_materials_df.to_excel(w, sheet_name='cream_data')
            final_times.to_excel(w, sheet_name='final_times')
        writer._save()
        return output.getvalue()

    df_xlsx = to_excel()
    st.download_button(label='📥 Скачать план в Excel', data=df_xlsx, file_name='Safia_Plan.xlsx')



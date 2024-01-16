import re

import openpyxl
import pandas as pd
from pptx import Presentation
from decimal import Decimal
# 1. Прочитать Excel с помощью pandas
s_n = 'Преза 20'#########################################################################################################
df = pd.read_excel('Справка по субъекту.xlsx', sheet_name=s_n)#Данные для презентации.xlsx###############################
df.fillna("", inplace=True)
#print(df['replace_value'])

# Проверка на наличие нужных столбцов
if 'metka' not in df.columns or 'replace_value' not in df.columns:
    raise ValueError('Не найдены столбцы "metka" или "replace_value"')

def format_value(value):
    if isinstance(value, (float, int)):
        value = Decimal(str(round(value, 2)))  # Преобразовать значение в Decimal для точности
        if value % 1 == 0:  # проверка, является ли число целым
            return str(int(value))
        else:
            formatted_value = str(value).rstrip('0').rstrip('.')
            return formatted_value.replace('.', ',')
    else:
        return str(value)

replace_dict = {metka: format_value(replace_value) for metka, replace_value in zip(df['metka'], df['replace_value'])}

# 2. Открыть презентацию и произвести замены
presentation = Presentation('Заготовка Слайд_НММО.pptx')#Заготовка_Презентация_НММО_расширенная2023.pptx#################

def replace_text_in_runs(runs):
    for run in runs:
        for metka, replace_value in replace_dict.items():
            run.text = run.text.replace(metka, replace_value)

def process_shapes_for_replacement(shapes):
    for shape in shapes:
        if shape.has_text_frame:
            paragraphs = shape.text_frame.paragraphs
            for paragraph in paragraphs:
                replace_text_in_runs(paragraph.runs)

        # Обработка таблиц
        elif shape.has_table:
            table = shape.table
            for row in table.rows:
                for cell in row.cells:
                    paragraphs = cell.text_frame.paragraphs
                    for paragraph in paragraphs:
                        replace_text_in_runs(paragraph.runs)

for slide in presentation.slides:
    # Обработка основного текста слайдов
    process_shapes_for_replacement(slide.shapes)

    # Обработка заметок лектора
    if slide.has_notes_slide:
        paragraphs = slide.notes_slide.notes_text_frame.paragraphs
        for paragraph in paragraphs:
            replace_text_in_runs(paragraph.runs)


presentation.save('измененная_презентация.pptx')

import os
import zipfile
import shutil
from openpyxl import load_workbook

def update_embedded_excel(pptx_name, data_file):
    temp_dir = 'temp_pptx_content'

    # 1. Разархивировать содержимое .pptx во временную папку
    with zipfile.ZipFile(pptx_name, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    # 2. Изменить встроенный Excel документ
    data_df = pd.read_excel(data_file, sheet_name=s_n)
    for index, row in data_df.iterrows():
        if str(row['A']).startswith("ppt/"):
            file_path = os.path.join(temp_dir, row['A'])
            # current_file_path = os.path.abspath(__file__)
            # current_directory = os.path.dirname(current_file_path)
            # # Создание полного пути к файлу
            # file_path = os.path.join(current_directory, 'temp_pptx_content', str(row['A']))
            if os.path.exists(file_path):
                with open(file_path, 'rb') as f:
                    embedded_df = pd.read_excel(f)

                rows_count = int(row['B'])
                cols_count = int(row['C'])

                # Выберите данные, которые вы хотите добавить в embedded_df
                new_data = data_df.iloc[index + 1: index + 1 + rows_count, :cols_count]
                new_data = new_data.reset_index(drop=True)

                # Объедините новые данные с embedded_df
                embedded_df = pd.concat([new_data, embedded_df], ignore_index=True)
                #new_columns = embedded_df.iloc[0, :cols_count].values
                #embedded_df.columns = list(new_columns) + list(embedded_df.columns[cols_count:])
                #embedded_df = embedded_df.drop(0).reset_index(drop=True)
                embedded_df = embedded_df.iloc[:rows_count, :cols_count]

                #with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                #    embedded_df.to_excel(writer, index=False)
                wb = load_workbook(file_path)
                ws = wb["Лист1"]
                for row_index, row in embedded_df.iterrows():
                    for col_index, value in enumerate(row):
                        # Прибавляем 1, так как openpyxl индексирует с 1
                        ws.cell(row=row_index + 1, column=col_index + 1, value=value)
                print(wb)
                print(embedded_df)
                wb.save(file_path)


    # 3. Заархивировать содержимое временной папки
    with zipfile.ZipFile(pptx_name, 'w') as zip_ref:
        for foldername, subfolders, filenames in os.walk(temp_dir):
            for filename in filenames:
                file_path = os.path.join(foldername, filename)
                archive_path = file_path[len(temp_dir) + 1:]
                zip_ref.write(file_path, archive_path)

    # Удаляем временную папку
    shutil.rmtree(temp_dir)


# pptx_name = "измененная_презентация.pptx"
# data_file = "Данные для презентации.xlsx"
# update_embedded_excel(pptx_name, data_file)
#
# import win32com.client
# import time
# def update_powerpoint_charts(pptx_path):
#     # Создаем COM-объект для PowerPoint
#     powerpoint = win32com.client.Dispatch("PowerPoint.Application")
#
#     # Открываем презентацию
#     presentation = powerpoint.Presentations.Open(pptx_path)
#
#     for slide in presentation.Slides:
#         for shape in slide.Shapes:
#             if shape.HasChart:
#                 # Здесь мы пытаемся "обновить" диаграмму. В вашем случае это может и не привести к результату,
#                 # но это максимум, что можно сделать средствами COM-объекта PowerPoint
#                 shape.Chart.ChartData.Activate()
#                 time.sleep(0.5)
#                 shape.Chart.ChartData.BreakLink()
#                 time.sleep(0.5)
#                 shape.Chart.Refresh()
#
#                 print(f"Обработка диаграммы на слайде {slide.SlideNumber}, название диаграммы: {shape.Name}")
#
#     # Сохраняем и закрываем презентацию
#     presentation.Save()
#     presentation.Close()
#
#     # Закрываем PowerPoint
#     powerpoint.Quit()
#
# current_directory = os.path.dirname(os.path.abspath(__file__))
# pptx_path = os.path.join(current_directory, "измененная_презентация.pptx")
# update_powerpoint_charts(pptx_path)

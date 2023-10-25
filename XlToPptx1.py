import openpyxl
import pandas as pd
from pptx import Presentation
from decimal import Decimal
# 1. Прочитать Excel с помощью pandas
df = pd.read_excel('Данные для презентации.xlsx', sheet_name='01')
print(df['replace_value'])

# Проверка на наличие нужных столбцов
if 'metka' not in df.columns or 'replace_value' not in df.columns:
    raise ValueError('Не найдены столбцы "metka" или "replace_value"')

def format_value(value):
    if isinstance(value, (float, int)):
        value = Decimal(str(value))  # Преобразовать значение в Decimal для точности
        if value % 1 == 0:  # проверка, является ли число целым
            return str(int(value))
        else:
            formatted_value = str(value).rstrip('0').rstrip('.')
            return formatted_value.replace('.', ',')
    else:
        return str(value)

replace_dict = {metka: format_value(replace_value) for metka, replace_value in zip(df['metka'], df['replace_value'])}

# 2. Открыть презентацию и произвести замены
presentation = Presentation('Заготовка_Презентация_НММО_расширенная2023.pptx')

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

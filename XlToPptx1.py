import openpyxl
from pptx import Presentation
# 1. Прочитать Excel и создать словарь для замен
wb = openpyxl.load_workbook('Данные для презентации.xlsx')
sheet = wb.active

replace_dict = {}
for row in sheet.iter_rows(min_row=2, values_only=True):  # пропускаем заголовки
    metka, replace_value = row
    replace_dict[metka] = str(replace_value)

# 2. Открыть презентацию и произвести замены
presentation = Presentation('Заготовка_Презентация_НММО_расширенная2023.pptx')
def replace_text_in_runs(runs):
    for run in runs:
        for metka, replace_value in replace_dict.items():
            run.text = run.text.replace(metka, replace_value)

for slide in presentation.slides:
    # Обработка основного текста слайдов
    for shape in slide.shapes:
        if shape.has_text_frame:
            paragraphs = shape.text_frame.paragraphs
            for paragraph in paragraphs:
                replace_text_in_runs(paragraph.runs)

    # Обработка заметок лектора
    if slide.has_notes_slide:
        paragraphs = slide.notes_slide.notes_text_frame.paragraphs
        for paragraph in paragraphs:
            replace_text_in_runs(paragraph.runs)

presentation.save('измененная_презентация.pptx')
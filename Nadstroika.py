import os
import re
from os.path import getmtime
def newest(pattern, rank=-1, search_in_subfolders=0, full_path=0):
# возвращает самый новый файл соответствующий паттерну
#          /паттерн /ранг с конца                    /вывести полный путь папки
#                            /искать ли в папках или указать имя папки
    root_dir = '.'  # Настраиваем начальную директорию для поиска
    if isinstance(search_in_subfolders, str):
        root_dir = search_in_subfolders  # Если это путь, используем его
    elif search_in_subfolders:
        root_dir = '.'  # Если истина, ищем в текущей директории и всех поддиректориях

    file_paths = []
    for root, dirs, files in os.walk(root_dir):
        # Обходим файлы и проверяем совпадение с регулярным выражением
        for name in files:
            if re.match(pattern, name):
                file_paths.append(os.path.join(root, name))
        # Если поиск в поддиректориях не требуется, останавливаемся после первой итерации
        if not isinstance(search_in_subfolders, str) and not search_in_subfolders:
            break

    if not file_paths:
        return None

    # Сортируем файлы по времени изменения
    file_paths.sort(key=lambda x: getmtime(x), reverse=(rank < 0))
    adjusted_rank = abs(rank) - 1
    if adjusted_rank >= len(file_paths):
        return None

    # Выбираем файл по рангу и определяем, нужен полный путь или только имя
    file_path = file_paths[adjusted_rank]
    return os.path.abspath(file_path) if full_path else os.path.basename(file_path)

# Пример использования функции с сырой строкой для регулярного выражения
result = newest(r'.+\.bat', -1 , "venv/Scripts", 1)

import os
import pandas as pd

# Путь к папке с файлами XLSX
folder_path = r'C:\Users\kyana\Desktop\test\test'

# Список всех файлов XLSX в папке
file_list = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]

# Создание пустого DataFrame для хранения объединенных данных
merged_data = pd.DataFrame()

# Заданная структура столбцов
desired_columns = ['НомСч', 'ПорНом', 'ПорНомДФ', 'ВремяОпер', 'ДатаОпер', 'НазПлат', 'ПопМСКВремя',
                   'ПорНомБлок', 'СумДебета', 'СумКредита', 'ВидДок', 'ДатаДок', 'НомДок',
                   'НаимБП', 'БИКБП', 'ИННПП', 'КПППП', 'НаимПП', 'НомСЧПП',
                   'НомКорСЧ', 'Должность', 'Имя', 'Отчество', 'Фамилия']

# Проход по каждому файлу XLSX
for file_name in file_list:
    file_path = os.path.join(folder_path, file_name)

    # Чтение файла XLSX и получение данных с четвертой строки
    df = pd.read_excel(file_path, header=None, skiprows=3)

    # Проверка наличия столбцов в файле
    existing_columns = set(df.iloc[0].tolist())
    columns_to_include = list(set(desired_columns) & existing_columns)

    # Проверка наличия хотя бы одного столбца для включения
    if columns_to_include:
        # Переупорядочивание и включение только существующих столбцов
        df = df[columns_to_include]

        # Объединение данных с помощью метода concat
        merged_data = pd.concat([merged_data, df], ignore_index=True)

# Проверка размера объединенных данных
if merged_data.shape[0] > 900000:
    # Разделение данных на части с ограничением в 900000 строк
    num_chunks = merged_data.shape[0] // 900000 + 1

    for i in range(num_chunks):
        start_idx = i * 900000
        end_idx = (i + 1) * 900000

        chunk_df = merged_data[start_idx:end_idx]

        # Создание нового файла XLSX на основе части данных
        output_file_name = f"merg{i+1}.xlsx"
        output_file_path = os.path.join(folder_path, output_file_name)
        chunk_df.to_excel(output_file_path, index=False)
else:
    # Создание нового файла XLSX на основе объединенных данных
    output_file_path = os.path.join(folder_path, "merged.xlsx")
    merged_data.to_excel(output_file_path, index=False)

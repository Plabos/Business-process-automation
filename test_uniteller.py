import pandas as pd
import os


def get_uni():
    """
    Функция для чтения файла uni, очистки колонки AuthCode от лишнего символа 
    и получения необходимых для дальнейшей работы столбцов
    """
    
    # Чтения файла с необходимой кодировкой
    df = pd.read_csv('uni.csv', sep=';', encoding='cp1251')
    
    # Убираем символ, который может встречаться в колонке AuthCode
    df['AuthCode'] = df.AuthCode.apply(lambda auth_code: auth_code.replace("’", ""))
    
    # Оставляем нужные далее колонки
    df = df[['OrderID', 'AuthCode']]
    
    return df


def get_statements():
    """
    Функция для получения выписок и объединения их в один файл
    """
    
    # Получаем путь к текущей директории(здесь лежат файлы выписок, uni и сам скрипт)
    way = os.getcwd()
    
    # При необходимости, можно отдельно указать путь до папки, где лежат выписки
    # в остальном предполагается, что uni и скрипт лежат в одной папке
    # way = '~/'
    
    
    # Собираем все названия файлов, где есть слово "Выписка"
    names_of_statements = [i for i in os.listdir(way) if i.startswith('Выписка')]
    
    # Создаем пустой датафрейм, в который далее запишем выписки
    df = pd.DataFrame()
    
    # По очереди читаем все файлы выписок и добавляем их созданный ранее датафрейм
    for file in names_of_statements:
        statement = pd.read_excel(os.path.join(way, file))
        df = pd.concat([df, statement], ignore_index=True)
        
    return df


def adding_data(df_1, df_2):
    """
    Функция, объединяет датафрейм с выписками с недостающей информацией и рассчитывает недостающие данные
    """
    from datetime import datetime
    
    # Объединение двух фреймов.
    df = df_1.merge(df_2, how='left', left_on='Код авторизации', right_on='AuthCode')
    
    # Комиссия
    commission = 0.0125
    
    # Берем значение времени на момент создания отчета и в нужном формате добавляем 
    df['Время обработки'] = datetime.now().strftime("%d.%m.%Y %H:%M")
    
    # Высчитываем комиссию и сумму к переводу с учетом комиссии
    df['Комиссия'] = round(df['Сумма операции'] * commission, 2)
    df['Сумма к переводу'] = round(df['Сумма операции'] * (1 - commission), 2)
    
    return df


def set_column_width(sheet, number_to_letter):
    """
    Функция, задающая ширину столбцов
    """
    
    dims = {}
    for row in sheet.rows:
        for cell in row:
            if cell.value:
                dims[cell.column] = max((dims.get(cell.column, 0), len(str(cell.value))))
                
    for col, value in dims.items():
        sheet.column_dimensions[number_to_letter[col]].width = value * 1.2

        
def set_border(sheet, cell_range):
    """
    Функция, "рисующая" границы ячеек, в заданном диапазоне
    """
    from openpyxl.styles import Border, Side
    
    thin = Side(border_style="thin", color="000000")
    for row in sheet[cell_range]:
        for cell in row:
            cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)

            
def create_report(df_1):
    """
    Функция, которая оставляет нужные столбцы, записывает таблицу в файл (.xlsx)
    и форматирует внешний вид таблицы
    """
    
    # Библиотека, для более тонкой работы с Excel
    import openpyxl
    from openpyxl.utils.dataframe import dataframe_to_rows
    from openpyxl.styles import NamedStyle, Font, Alignment
    
    # Оставляем необходимые столбцы и задаем нужный нам порядок
    df = df_1[['Номер устройства', 'Дата операции', 'Дата обработки', 'Сумма операции',
       'Торговая уступка', 'К перечислению', 'РРН', 'Тип операции',
       'Код авторизации', 'Время обработки', 'OrderID', 'Комиссия', 
        'Сумма к переводу', 'Номер платежного поручения', 'Дата платежного поручения']]
    
    # Делаем собственный стиль, который позднее используем для заголовка таблицы
    header = NamedStyle(name="header")
    header.font = Font(bold=True)
    header.alignment = Alignment(horizontal="center", vertical="top")
    
    # Создаю книгу и удаляю первый лист, который появляется автоматически при создании
    book = openpyxl.Workbook()
    book.remove(book.active)
    
    # Создаем лист, на котором будет таблица
    sheet_1 = book.create_sheet("Расчеты")
    
    # записываем полученную таблицу на лист
    for r in dataframe_to_rows(df, index=False, header=True):
        sheet_1.append(r)
    
    # форматируем заголовок    
    header_row = sheet_1[1]
    for cell in header_row:
        cell.style = header
        
    # Соответствие номера столбца буквам в заголовке для Excel
    number_to_letter = {1: 'A',   2: 'B',  3: 'C',  4: 'D', 
                        5: 'E',   6: 'F',  7: 'G',  8: 'H', 
                        9: 'I',  10: 'J', 11: 'K', 12: 'L', 
                        13: 'M', 14: 'N', 15: 'O', 16: 'P'}
    
    # Устанавливаем удобую для чтения ширину столбца
    set_column_width(sheet_1, number_to_letter)
    
    # Делаем границы для всех ячеек
    rows_num, col_num = df.shape
    cell_range_str = f'A1:{number_to_letter[col_num]}{rows_num + 1}'
    set_border(sheet_1, cell_range_str)
    
    # Сохраняем все изменения в файл (.xlsx)
    book.save("Отчёт.xlsx")
    
# Получаем файл uni
uni = get_uni()

# Получае все выписки в единый фрейм
bank_statements = get_statements()

# Дополняем данные необходимыми столбцами
augmented_data = adding_data(bank_statements, uni)

# Выводим полученные данные и приводим к нужному нам виду
create_report(augmented_data)
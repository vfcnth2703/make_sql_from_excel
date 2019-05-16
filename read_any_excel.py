import os
import sys
import xlrd
from transliterate import translit


def is_digit(string):
    """
        Проверяет является ли строка числом.
        На самом деле встроенная функция str.isdigit()
        не работает с float
    """
    if string.isdigit():
        return True
    else:
        try:
            float(string)
            return True
        except ValueError:
            return False


def add_space(val):
    """
        Делает заданный отступ
    """
    return ' ' * val


def collect_files(*args):
    """
        Делает список файлов с заданными расширения в
        текущей директории.
    """
    args = tuple(['.' + item for item in args])
    home_dir = os.getcwd()
    file_list = []
    for file in sorted(os.listdir(home_dir)):
        if file.endswith(args):
            file_list.append(home_dir + '\\' + file)
    return file_list


def read_excel_file(file_name, sheet_index=0):
    """
        Делаем список из данных excel файла
    """
    data = []
    rd = xlrd.open_workbook(file_name)
    sheet = rd.sheet_by_index(sheet_index)
    for rownum in range(sheet.nrows):
        row = sheet.row_values(rownum)
        data.append(row)
    return data


def trans(arg):
    return translit('_'.join(str(arg).split()).title(), 'ru', reversed=True)


def convert_num(arg):
    """
        Конвертируем все значения к строке
    """
    if is_digit(str(arg)):
        return str(float(arg))
    else:
        return str(arg)


def convert_header(data):
    """
        Формируем заголовок (список полей)
    """
    header = data[0]
    new_header = []
    for i, item in enumerate(header):
        if len(item) != 0:
            new_header.append(item)
        else:
            new_header.append(f'FLD_{str(i).zfill(4)}')
    new_header = '[' + '], ['.join(new_header) + ']'
    return(new_header)


def write_sql(header, body, file_name):
    """
        Пишем результирующий файл sql
    """
    with open(file_name, 'w') as f:
        f.write(f'WITH R ([ID], {header})\n AS (\n')
        for id, item in enumerate(body, start=1):
            line = '\', \''.join(map(convert_num, item))
            f.write(f'{add_space(4)}select {id},\'{line}\'\n')
            if item != body[-1]:
                f.write(f'{add_space(6)}union all\n')
        f.write(')\n')
        f.write(f'{add_space(4)}Select * from R\n\n')


def set_ext_to_file(file, ext):
    """
        Меняем расширение файла
    """
    return f'{os.path.splitext(file)[0]}.{ext}'


def main():
    for excel_file in collect_files('xls', 'xlsx'):
        data = read_excel_file(excel_file)
        sql_file = set_ext_to_file(excel_file, 'sql')
        header = convert_header(data)
        body = data[1:]
        write_sql(header, body, sql_file)


if __name__ == '__main__':
    main()

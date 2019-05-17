from functions import *

def main():
    for excel_file in collect_files('xls', 'xlsx'):
        data = read_excel_file(excel_file)
        sql_file = set_ext_to_file(excel_file, 'sql')
        header = data[0]
        body = data[1:]
        header = convert_header(header)
        write_sql(header, body, sql_file)


if __name__ == '__main__':
    main()

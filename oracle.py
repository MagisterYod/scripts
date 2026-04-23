import oracledb
import xlwt
import os
from SqlAndMail import SqlAndMail

lib_dir = "C:\\app\\product\\instantclient_12_2"
oracledb.init_oracle_client(lib_dir=lib_dir)
connection = oracledb.connect(user=os.getenv('USER'), password=os.getenv('PASSWORD_BD'), dsn=os.getenv('DNS'))
cursor = connection.cursor()

receiver_email = [
    'OS2@polymetal.ru',
    'kuznecovia@polymetal.ru'
]

result_table = {}
sql = f"""
select
    *
from
    ZSU.WELLMETERARCHIVE wa
where
    WA.TERMINALTIME between to_date('{SqlAndMail.p_date(1)} 00:00:00', 'DD.MM.YYYY HH24:MI:SS')
    and to_date('{SqlAndMail.p_date(1)} 23:59:00', 'DD.MM.YYYY HH24:MI:SS')
and
    WA.TRUCK_KEY = 326
order by key asc
"""
for r in cursor.execute(sql):
    time_table = r[9].strftime('%H')
    volume = r[6]
    key = r[0]
    if time_table in result_table:
        tm = result_table.get(time_table)
        if tm[0] < key:
            tm[2] = volume
            tm[3] = tm[2] - tm[1]
        else:
            tm[0] = key
            tm[1] = volume
        result_table[time_table] = tm
    else:
        result_table[time_table] = [key, volume, 0, 0]

work_book = xlwt.Workbook(encoding='utf-8')
work_sheet = work_book.add_sheet('Отчет по воде', cell_overwrite_ok=True)
work_sheet.row(0).height_mismatch = True
work_sheet.row(0).height = 256 * 2
row_num = 2
# Инициализация стиля заголовка
style_head = xlwt.XFStyle()
# Настройка шрифта заголовка
style_head.font.name = 'Calibre'
style_head.font.height = 20 * 10
style_head.font.bold = True
# Настройка границ заголовка
style_head.borders = xlwt.Borders()
style_head.borders.left = 2
style_head.borders.right = 2
style_head.borders.top = 2
style_head.borders.bottom = 2
# Настройка выравнивания заголовка
style_head.alignment.horz = 0x02
style_head.alignment.vert = 0x01

work_sheet.write_merge(
    r1=1, r2=1, c1=0, c2=3, label=f'Отчет по питьевой воде за {SqlAndMail.p_date(1)}', style=style_head
)
# Настройка цвета заголовка
pattern = xlwt.Pattern()
pattern.pattern = xlwt.Pattern.SOLID_PATTERN
pattern.pattern_fore_colour = 44
style_head.pattern = pattern

columns = ['Час', 'Начало по счетчику', 'Конец по счетчику', 'Расход за период']
for col_num in range(len(columns)):
    if col_num == 0:
        work_sheet.col(0).width = 7 * 256
        work_sheet.write(row_num, col_num, columns[col_num], style_head)
    else:
        work_sheet.col(col_num).width = 23 * 256
        work_sheet.write(row_num, col_num, columns[col_num], style_head)
row_num += 1
# Инициализация стиля устроенного
style_insert = xlwt.XFStyle()
# Настройка шрифта
style_insert.font.name = 'Calibri'
style_insert.font.height = 20 * 10
style_insert.font.bold = False
# style_insert.font.colour_index = 3
# Настройка границ устроенного
style_insert.borders.left = 1
style_insert.borders.right = 1
style_insert.borders.top = 1
style_insert.borders.bottom = 1
# Настройка выравнивания устроенного
style_insert.alignment.horz = 0x02
style_insert.alignment.vert = 0x01

for index, value in result_table.items():
    work_sheet.write(row_num, 0, index, style_insert)
    value.pop(0)
    for col_value in range(len(value)):
        work_sheet.write(row_num, col_value + 1, round(value[col_value], 2), style_insert)
    row_num += 1

work_book.save(f'{SqlAndMail.p_date(1)}.xls')

SqlAndMail.send_emails(
    receiver_email=receiver_email,
    body=f'Добрый день!\n\nЭто автоматическое письмо! Отвечать на него не надо.',
    subject=f'Ежесменный отчет по питьевой воде за {SqlAndMail.p_date(1)}',
    file_name=f'{SqlAndMail.p_date(1)}.xls'
)

os.remove(f'{SqlAndMail.p_date(1)}.xls')
cursor.close()
print('Done water!')

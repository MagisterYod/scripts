import xlwt
from xls2xlsx import XLS2XLSX
import os
import pandas
import datetime
import smtplib
from time import sleep
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv
from SqlAndMail import SqlAndMail

dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
if os.path.exists(dotenv_path):
    load_dotenv(dotenv_path)

dt = datetime.datetime.now().strftime('%Y_%m_%d')
dt_time = datetime.datetime.now().strftime('%H_%M')
# down_path = '\\\\172.23.100.9\Stolovaya\СБ\Выгрузки\тест'
down_path = '\\\\172.23.100.9\Stolovaya\СБ\Выгрузки'
# up_path = '\\\\172.23.100.9\Stolovaya\СБ\Выгрузки'
up_path = '\\\\172.23.100.9\Stolovaya\СБ'
back_path = '\\\\172.23.100.9\Stolovaya\СБ\Выгрузки\BackUp'
# smtp_server = '172.23.100.20'
# smtp_port = 25
# sender_email = 'dr_failed@polymetal.ru'
receiver_email = 'kuznecovia@polymetal.ru'

result = {
    'zsu': [],
    'sgk': [],
    'kpm': [],
    'uf': [],
    'ml': [],
    'mr': [],
    'psu': []
}


def add_list_dict(org, add_df):
    list_dict = result[org]
    list_dict.append(add_df)
    result[org] = list_dict


def parse_df_exl(path) -> dict:

    df = pandas.read_excel(f'{down_path}\\{path}', header=None)
    if len(df) == 0:
        df = list(df)
        df.pop(0)
        var_list = 1
        while var_list <= 4:
            df.append(' ')
            var_list += 1
        if str(df[4]).startswith("U"):
            df[4] = ' '
        if df[5] == 'Золото Северного Урала':
            df[5] = 'ЗСУ'
            add_list_dict('zsu', df)
        if df[5] == 'Саумская Горнорудная Компания':
            df[5] = 'СГК'
            add_list_dict('sgk', df)
        if df[5] == 'Краснотурьинск-Полиметалл':
            df[5] = 'КПМ'
            add_list_dict('kpm', df)
        if df[5] == 'Уральский филиал Полиметалл УК' or df[5] == 'Уральский филиал':
            df[5] = 'УФ'
            add_list_dict('uf', df)
        if df[5] == 'ООО Минераллаб':
            df[5] = 'МЛ'
            add_list_dict('ml', df)
        if df[5] == 'Минерал Ресурс':
            df[5] = 'МР'
            add_list_dict('mr', df)
        if df[5] == 'ООО Полиметаллы Северного Урала':
            df[5] = 'ПСУ'
            add_list_dict('psu', df)
    else:
        for i, r in df.iterrows():
            row = list(r)
            row.pop(0)
            var_list = 1
            while var_list <= 4:
                row.append(' ')
                var_list += 1
            if str(row[4]).startswith("U"):
                row[4] = ' '
            if row[5] == 'Золото Северного Урала':
                row[5] = 'ЗСУ'
                add_list_dict('zsu', row)
                continue
            if row[5] == 'Саумская Горнорудная Компания':
                row[5] = 'СГК'
                add_list_dict('sgk', row)
                continue
            if row[5] == 'Краснотурьинск-Полиметалл':
                row[5] = 'КПМ'
                add_list_dict('kpm', row)
                continue
            if row[5] == 'Уральский филиал Полиметалл УК' or row[5] == 'Уральский филиал':
                row[5] = 'УФ'
                add_list_dict('uf', row)
                continue
            if row[5] == 'ООО Минераллаб':
                row[5] = 'МЛ'
                add_list_dict('ml', row)
                continue
            if row[5] == 'Минерал Ресурс':
                row[5] = 'МР'
                add_list_dict('mr', row)
            if row[5] == 'ООО Полиметаллы Северного Урала':
                row[5] = 'ПСУ'
                add_list_dict('psu', row)
    return result


def insert_work_sheet(df, organization):

    def inner_insert_def(data_frame, style):
        for index_cell in range(len(data_frame) - 1):
            if index_cell == 0:
                work_sheet.write(row_num, index_cell, index_row + 1, style)
                work_sheet.write(row_num, index_cell + 1, df[index_row][index_cell], style)
            else:
                work_sheet.write(row_num, index_cell + 1, df[index_row][index_cell], style)

    work_book = xlwt.Workbook(encoding='utf-8')
    work_sheet = work_book.add_sheet('Лист1', cell_overwrite_ok=True)
    work_sheet.row(0).height_mismatch = True
    work_sheet.row(0).height = 256 * 2
    row_num = 4
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
    # Настройка цвета заголовка
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = 44
    style_head.pattern = pattern
    columns = [
        '№П/П',
        'т. номер',
        'фамилия',
        'имя',
        'отчество',
        'ключ/карта',
        'Организация',
        'Дотация',
        'К удержанию',
        ''
    ]
    for col_num in range(len(columns)):
        if col_num == 0:
            work_sheet.col(0).width = 7 * 256
            work_sheet.write(row_num, col_num, columns[col_num], style_head)
        else:
            work_sheet.col(col_num).width = 23 * 256
            work_sheet.write(row_num, col_num, columns[col_num], style_head)
    # Инициализация стиля устроенного
    style_green = xlwt.XFStyle()
    # Настройка шрифта
    style_green.font.name = 'Calibri'
    style_green.font.height = 20 * 10
    style_green.font.bold = False
    style_green.font.colour_index = 3
    # Настройка границ устроенного
    style_green.borders.left = 1
    style_green.borders.right = 1
    style_green.borders.top = 1
    style_green.borders.bottom = 1
    # Настройка выравнивания устроенного
    style_green.alignment.horz = 0x02
    style_green.alignment.vert = 0x01
    # Инициализация стиля уволенного
    style_red = xlwt.XFStyle()
    # Настройка шрифта уволенного
    style_red.font.name = 'Calibri'
    style_red.font.height = 20 * 10
    style_red.font.bold = False
    style_red.font.colour_index = 2
    # Настройка границ уволенного
    style_red.borders.left = 1
    style_red.borders.right = 1
    style_red.borders.top = 1
    style_red.borders.bottom = 1
    # Настройка выравнивания уволенного
    style_red.alignment.horz = 0x02
    style_red.alignment.vert = 0x01
    for index_row in range(len(df)):
        row_num += 1
        if df[index_row][4] != " ":
            inner_insert_def(df[index_row], style_green)
        else:
            inner_insert_def(df[index_row], style_red)
    work_book.save(f'{down_path}\card_{organization}_{dt}.xls')
    xls = XLS2XLSX(f'{down_path}\card_{organization}_{dt}.xls')
    xls.to_xlsx(f'{up_path}\card_{organization}_{dt}.xlsx')
    os.remove(f'{down_path}\card_{organization}_{dt}.xls')


def back_rename_file(down_path, filename):
    os.rename(f"{down_path}\\{filename}", f'{back_path}\\{filename.split(".")[0]}_{dt}_{dt_time}.xls')


# def send_error(e, filename):
#     with smtplib.SMTP(smtp_server, smtp_port) as server:
#         server.starttls()
#         server.login(os.getenv("LOGIN"), os.getenv("PASSWORD"), initial_response_ok=True)
#         body = f'Не добрый день!\n\nОшибка: {e}\n\nФайл: {filename}\n\n'
#         message = MIMEMultipart()
#         message["Subject"] = f'{datetime.datetime.now().strftime("%d.%m.%Y")} {e.__class__}'
#         message["From"] = sender_email
#         message["To"] = receiver_email
#         message_text = MIMEText(body.encode('utf-8'), 'plain', 'utf-8')
#         message.attach(message_text)
#         server.ehlo()  # Может быть опущено
#         server.sendmail(sender_email, receiver_email, message.as_string())
#         server.quit()


for filename in os.listdir(f'{down_path}'):
    if os.path.isfile(f'{down_path}\\{filename}'):
        try:
            parse_df_exl(filename)
            back_rename_file(down_path, filename)
            print("Ok!")
        except Exception as e:
            SqlAndMail.send_emails(
                receiver_email,
                body='Не добрый день!\n\nОшибка: {e}\n\nФайл: {filename}\n\n',
                subject=f'{datetime.datetime.now().strftime("%d.%m.%Y")} {e.__class__}'
            )
            # send_error(e, filename)
            print("End with error")
            sleep(4)
            continue
        sleep(1)


for i in result:
    if result[i]:
        print(result[i])
        insert_work_sheet(result[i], i)

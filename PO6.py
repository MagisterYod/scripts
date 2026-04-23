import os
import openpyxl
from SqlAndMail import SqlAndMail


receiver_email = [
    'TCittcerDA@polymetal.ru',
    'UstsovAV@polymetal.ru',
    'Kotov@polymetal.ru',
    'RootSA@polymetal.ru',
    'kuznecovia@polymetal.ru'
]

months = {
    '01': 'Январь',
    '02': 'Февраль',
    '03': 'Март',
    '04': 'Апрель',
    '05': 'Май',
    '06': 'Июнь',
    '07': 'Июль',
    '08': 'Август',
    '09': 'Сентябрь',
    '10': 'Октябрь',
    '11': 'Ноябрь',
    '12': 'Декабрь'
}

result = {}


def insert_row(dict_res, num, _row):
    if dict_res.get(num) is None:
        result[num] = [_row]
    else:
        item_dict = dict_res.get(num)
        item_dict.append(_row)
        result[num] = item_dict


def insert_technics_result_month(work_sheet, row_tech, df):
    work_sheet[f'F{row_tech}'] = df[3]  # 6 ДВС
    work_sheet[f'H{row_tech}'] = df[4]  # 8 пробег общий
    work_sheet[f'I{row_tech}'] = df[5]  # 9 пробег с грузом
    work_sheet[f'K{row_tech}'] = df[6]  # 11 объем работ план
    work_sheet[f'L{row_tech}'] = df[7]  # 12 объем работ факт
    work_sheet[f'M{row_tech}'] = df[8]  # 13 план ТО
    work_sheet[f'N{row_tech}'] = df[9]  # 14 факт ТО
    work_sheet[f'O{row_tech}'] = df[10]  # 16 планТР
    work_sheet[f'P{row_tech}'] = df[11]  # 16 факт ТР
    work_sheet[f'Q{row_tech}'] = df[12]  # 18 план КР
    work_sheet[f'R{row_tech}'] = df[13]  # 18 фатк КР
    work_sheet[f'S{row_tech}'] = df[14]  # 20 план регламент
    work_sheet[f'T{row_tech}'] = df[15]  # 20 факт регламент
    work_sheet[f'U{row_tech}'] = df[16]  # 22 план обед
    work_sheet[f'V{row_tech}'] = df[17]  # 22 факт обед
    work_sheet[f'X{row_tech}'] = df[18]  # 24 факт прием/передача
    work_sheet[f'Y{row_tech}'] = df[19]  # 26 план забой
    work_sheet[f'Z{row_tech}'] = df[20]  # 26 факт забой
    work_sheet[f'AA{row_tech}'] = df[21]  # 28 план БВР
    work_sheet[f'AB{row_tech}'] = df[22]  # 28 факт БВР
    work_sheet[f'AC{row_tech}'] = df[23]  # 29 ДВС
    work_sheet[f'AD{row_tech}'] = df[24]  # 30 трансмиссия
    work_sheet[f'AE{row_tech}'] = df[25]  # 31 ходовая
    work_sheet[f'AF{row_tech}'] = df[26]  # 32 навесное
    work_sheet[f'AG{row_tech}'] = df[27]  # 33 электро
    work_sheet[f'AH{row_tech}'] = df[28]  # 34 гибравлика
    work_sheet[f'AI{row_tech}'] = df[29]  # 35 прочие
    work_sheet[f'AJ{row_tech}'] = df[30]  # 36 автошины
    work_sheet[f'AK{row_tech}'] = df[31]  # 37 фронт работ
    work_sheet[f'AL{row_tech}'] = df[32]  # 38 зап.части
    work_sheet[f'AM{row_tech}'] = df[33]  # 39 ГСМ
    work_sheet[f'AN{row_tech}'] = df[34]  # 40 персонал
    work_sheet[f'AO{row_tech}'] = df[35]  # 41 метеоусловия
    work_sheet[f'AP{row_tech}'] = df[36]  # 42 прочие
    work_sheet[f'AT{row_tech}'] = float(df[37])  # 46 фонд времени


for num_month in range(1, int(SqlAndMail.p_date_sep(3)[1]) + 1):
    if num_month < 10:
        num_month = f'0{num_month}'
    else:
        num_month = num_month
    for row in (SqlAndMail.cursor_data(
            query='ZSU.GET_BUDGET_RESULT_MONTH',
            date=str(f'01.{num_month}.{SqlAndMail.p_date_sep(3)[2]}')
    ).getvalue().fetchall()
    ):
        if int(row[1]) == 45:
            insert_row(result, num_month, row)
        if int(row[1]) == 28:
            insert_row(result, num_month, row)
        if int(row[1]) == 31:
            insert_row(result, num_month, row)
        if int(row[1]) == 1:
            insert_row(result, num_month, row)
        if int(row[1]) == 272:
            insert_row(result, num_month, row)
        if int(row[1]) == 9:
            insert_row(result, num_month, row)
        if int(row[1]) == 10:
            insert_row(result, num_month, row)
        if int(row[1]) == 186:
            insert_row(result, num_month, row)
        if int(row[1]) == 4:
            insert_row(result, num_month, row)
        if int(row[1]) == 15:
            insert_row(result, num_month, row)
        if int(row[1]) == 58:
            insert_row(result, num_month, row)
        if int(row[1]) == 43:
            insert_row(result, num_month, row)
        if int(row[1]) == 183:
            insert_row(result, num_month, row)
        if int(row[1]) == 283:
            insert_row(result, num_month, row)
        if int(row[1]) == 355:
            insert_row(result, num_month, row)
        if int(row[1]) == 284:
            insert_row(result, num_month, row)
        if int(row[1]) == 285:
            insert_row(result, num_month, row)
        if int(row[1]) == 286:
            insert_row(result, num_month, row)

wb = openpyxl.load_workbook(f'C:\\Users\\it_kuznecovia\\Documents\\ProjectsDjango\\webASD\\po6\\static\\PO6.xlsx')
for item, value in result.items():
    ws = wb[str(int(item))]
    ws['A9'] = f'за {months.get(str(item))} {SqlAndMail.p_date_sep(3)[2]} г.'
    for row in value:
        if row[1] == 45:
            insert_technics_result_month(ws, 23, row)
        if row[1] == 28:
            insert_technics_result_month(ws, 26, row)
        if row[1] == 31:
            insert_technics_result_month(ws, 27, row)
        if row[1] == 1:
            insert_technics_result_month(ws, 43, row)
        if row[1] == 272:
            insert_technics_result_month(ws, 45, row)
        if row[1] == 9:
            insert_technics_result_month(ws, 47, row)
        if row[1] == 10:
            insert_technics_result_month(ws, 48, row)
        if row[1] == 186:
            insert_technics_result_month(ws, 50, row)
        if row[1] == 4:
            insert_technics_result_month(ws, 53, row)
        if row[1] == 15:
            insert_technics_result_month(ws, 54, row)
        if row[1] == 58:
            insert_technics_result_month(ws, 56, row)
        if row[1] == 43:
            insert_technics_result_month(ws, 57, row)
        if row[1] == 183:
            insert_technics_result_month(ws, 58, row)
        if row[1] == 283:
            insert_technics_result_month(ws, 59, row)
        if row[1] == 355:
            insert_technics_result_month(ws, 60, row)
        if row[1] == 284:
            insert_technics_result_month(ws, 61, row)
        if row[1] == 285:
            insert_technics_result_month(ws, 62, row)
        if row[1] == 286:
            insert_technics_result_month(ws, 63, row)

wb.save(f'PO6_years_{SqlAndMail.p_date_sep(3)[1]}.{SqlAndMail.p_date_sep(3)[2]}.xlsx')

SqlAndMail.send_emails(
    receiver_email=receiver_email,
    body='Добрый день!\n\nЭто автоматическое рассылка! Отвечать на него не надо.',
    subject=f'Месячный отчет ПО-6 на дату {SqlAndMail.now()}',
    file_name=f'PO6_years_{SqlAndMail.p_date_sep(3)[1]}.{SqlAndMail.p_date_sep(3)[2]}.xlsx'
)

os.remove(f'PO6_years_{SqlAndMail.p_date_sep(3)[1]}.{SqlAndMail.p_date_sep(3)[2]}.xlsx')
print('Done PO6! ')

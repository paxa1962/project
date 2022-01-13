import xlwings as xw

wb = xw.Book('workbook.xlsx')
sheet = wb.sheets['miir_pptk_stat_col']
#print(sheet.range('I:I')[1:10].value) #Весь столбец с 1 по 10 строку

rng = sheet.range((2, 'G'), (7027, 'I')).value #Выбираем диапазон
for row in rng:
    if row[0] == 'CHANGEDATE':
        row[2] = 'Дата изменения'
    elif row[0] == 'VERDATE':
        row[2] = 'Дата верификации'
    elif row[0] == 'ISDELETED':
        row[2] = 'Флаг удаления'
    elif row[0] == 'ID':
        row[2] = 'Идентификатор'
    elif row[0] == 'SHORTNAME':
        row[2] = 'Короткое наименование или код'
    elif row[0] == 'ISSYSTEM':
        row[2] = 'Флаг системного атрибута'
    elif row[0] == 'CREATEDATE':
        row[2] = 'Дата создания'
    elif row[0] == 'CODENAME':
        row[2] = 'Кодовое имя'
    elif row[0] == 'SCHEMA':
        row[2] = 'Кодовое имя'
    elif row[0] == 'CLASSID':
        row[2] = 'Идентификатор класса'
    elif row[0] == 'DETAILS':
        row[2] = 'Коментарии'

sheet.range((2, 'G')).value = rng #записываем исправленный вариант

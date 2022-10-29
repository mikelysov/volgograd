import pandas as pd
import xlwt
import sys ; sys.setrecursionlimit(sys.getrecursionlimit() * 5)

sprotr = pd.read_excel('./inbox/Справочник. Отрасли и подотрасли.xlsx', header=1)
sprotr = sprotr.rename(columns={'Наименование подотрасли':'Подотрасль'})
sprotr.insert(0, 'otr', [s for s in range(len(sprotr))])

sprtech = pd.read_excel('./inbox/Справочник. Технологии.xlsx', header=1)
sprtech = sprtech.rename(columns={' 1 уровень':'Технология (1 уровень)', '2 уровень':'Технология (2 уровень)', '3 уровень (уровень тегирования участников)':'Технология (3 уровень)'})
sprtech.insert(0, 'tech', [s for s in range(len(sprtech))])

otr = pd.read_excel('./inbox/3. Отрасли.xlsx', header=0)
print(len(otr))
otr = otr.merge(sprotr, how='left', on=['Отрасль', 'Подотрасль'])
print(len(otr))

tech = pd.read_excel('./inbox/4. Технологии.xlsx', header=0)
print(len(tech))
tech = tech.merge(sprtech, how='left', on=['Технология (1 уровень)', 'Технология (2 уровень)', 'Технология (3 уровень)'])
print(len(tech))

org = pd.read_excel('./inbox/1. Компании.xlsx')
print(len(org))
org = org.merge(sprotr, how='left', on=['Отрасль', 'Подотрасль'])
org = org.merge(sprtech, how='left', on=['Технология (1 уровень)', 'Технология (2 уровень)', 'Технология (3 уровень)'])
print(len(org))
org.head(1)

prod = pd.read_excel('./inbox/2. Продукты_new.xlsx', header=1)
prod = prod.rename(columns={'global_id (продукта)':'global_id'})
prod.head(1)

result = []
for id, prow in prod.iterrows():
    found = otr[otr['Создаваемые продукты']==prow.global_id]
    if len(found) > 0:
        prodotr = found.otr.tolist()[0]
        if len(found) > 1:
            result.append('для продукта ' + str(prow.global_id) + ' найдено больше одной отрасли')
    else:
        prodotr = 9999

    found = tech[tech['Создаваемые продукты']==prow.global_id]
    if len(found) > 0:
        prodtech = found.tech.tolist()[0]
        if len(found) > 1:
            result.append('для продукта ' + str(prow.global_id) + ' найдено больше одной технологии')
    else:
        prodtech = 9999
    
    forg = org[org.global_id==prow.Компания]
    if len(forg) == 0:
        result.append('для продукта ' + str(prow.global_id) + ' не найдена организация')
        continue

    orgotr = forg.otr.tolist()[0]
    orgtech = forg.tech.tolist()[0]
    
    if prodotr!=orgotr:
        result.append('для продукта ' + str(prow.global_id) + ' для компании ' + str(prow.Компания) + ' не совпадают отрасли')
    if prodtech!=orgtech:
        result.append('у продукта ' + str(prow.global_id) + ' и компании ' + str(prow.Компания) + ' не совпадают технологии')


book = xlwt.Workbook(encoding='utf-8')
sheet = book.add_sheet('Ошибки анализа')#Errors while analysing organizations and products
i = 0
for s in result:
    sheet.write(i, 0, s);
    i += 1
book.save('./outbox/errors.xls')
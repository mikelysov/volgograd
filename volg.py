import numpy as np
import pandas as pd
import xlwt
import os
from keras.models import load_model
from keras.preprocessing.text import Tokenizer
from keras.utils import to_categorical
import sys ; sys.setrecursionlimit(sys.getrecursionlimit() * 5)

def load_set():
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

    return sprotr, sprtech, otr, tech, org, prod

def analyze_set(sprotr, sprtech, otr, tech, org, prod):
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
    
    return result

def save_t0_excel(result, filename):
    book = xlwt.Workbook(encoding='utf-8')
    sheet = book.add_sheet('Ошибки анализа')#Errors while analysing organizations and products
    i = 0
    for s in result:
        sheet.write(i, 0, s);
        i += 1
    book.save('./outbox/'+filename)

def analyze_desc(sprotr, sprtech, otr, tech, org, prod):
    VOCAB_SIZE = 20000
    WIN_SIZE = 10
    WIN_HOP = 5

    model_otr = load_model('./conf/otrasl.h5')
    #model_tech = load_model('./conf/techno.h5')    
    
    voc = sprotr.Отрасль.unique().tolist()
    voc += sprotr.Подотрасль.unique().tolist()
    voc += sprtech['Технология (1 уровень)'].unique().tolist()
    voc += sprtech['Технология (2 уровень)'].unique().tolist()
    voc += sprtech['Технология (3 уровень)'].unique().tolist()
    voc.remove(np.nan)

    tokenizer = Tokenizer(num_words=VOCAB_SIZE, filters='!"#$%&()*+,-–—./…:;<=>?@[\\]^_`{|}~«»\t\n\xa0\ufeff', lower=True, split=' ', oov_token='неизвестное_слово', char_level=False)
    tokenizer.fit_on_texts(voc)

    result = []
    for id, prow in prod.iterrows():
        found = otr[otr['Создаваемые продукты']==prow.global_id]
        x, y = [], []
        if len(found) > 0:
            otr_id = found.otr.tolist()[0]
            otr_lst = []
            otr_lst.append(found.Отрасль.tolist()[0] + ' ' + found.Подотрасль.tolist()[0])
            
            seq_list = tokenizer.texts_to_sequences(otr_lst)
            
            if len(seq_list[0]) < 10:
                seq_list[0] += [0 for i in range(10-len(seq_list[0]))]
                x = np.array(seq_list)
            else:
                vec = [seq_list[0][i:i+WIN_SIZE] for i in range(0, len(seq_list[0])-WIN_SIZE+1, WIN_HOP)]
                x = np.array(vec)
            
            y = [to_categorical(otr_id, 148)] * len(x)
            y = np.array(y)
        
            pred = model_otr.predict(x, verbose=0)
            otr_id_new = np.argmax(pred)
            if otr_id != otr_id_new:
                name1, name2 = found.Отрасль.tolist()[0], found.Подотрасль.tolist()[0]
                found = sprotr[sprotr.otr==otr_id_new]
                name3, name4 = found.Отрасль.tolist()[0], found.Подотрасль.tolist()[0]
                result.append('Для продукта %s отрасль "%s-%s" модель определила как "%s-%s"' %(prow.global_id, name1, name2, name3, name4))
    
    return result

def main():
    if not os.path.exists('./inbox'):
        print('Нет файлов для обработки')
        return
    
    fls = ['1. Компании.xlsx', '2. Продукты_new.xlsx', '3. Отрасли.xlsx', '4. Технологии.xlsx', 'Справочник. Отрасли и подотрасли.xlsx', 'Справочник. Технологии.xlsx']
    flag = True
    for s in fls:
        flag = flag and os.path.exists('./inbox/'+s)
    if not flag:
        print('Нет файлов для обработки')
        return

    if not os.path.exists('./outbox'):
        os.mkdir('./outbox')
    
    sprotr, sprtech, otr, tech, org, prod = load_set()
    
    #result = analyze_set(sprotr, sprtech, otr, tech, org, prod)
    #save_t0_excel(result, 'errors.xls')
    
    result = analyze_desc(sprotr, sprtech, otr, tech, org, prod)
    save_t0_excel(result, 'errors_desc.xls')

if __name__ == '__main__':
    main()

# Извлекать реплики с приветствием – где менеджер поздоровался.
# Извлекать реплики, где менеджер представил себя.
# Извлекать имя менеджера.
# Извлекать название компании.
# Извлекать реплики, где менеджер попрощался.
# Проверять требование к менеджеру: «В каждом диалоге обязательно необходимо поздороваться и попрощаться с клиентом»

import pandas as pd
import re
import pymorphy2

df = pd.read_csv(r"test_data.csv")
dfm = df[(df.role=="manager")]
morph = pymorphy2.MorphAnalyzer()


# ВЫДЕЛЕНИЕ НАЧАЛА И КОНЦА КАЖДОГО ДИАЛОГА

df_new = pd.DataFrame()
for dlg in df.dlg_id.unique():
    df_new = pd.concat([df_new, dfm[dfm.dlg_id==dlg].head(), dfm[dfm.dlg_id==dlg].tail()])

    
# РАСШИРЕНИЕ ТАБЛИЦЫ ДАННЫХ
    
company_pattern = r"((?<=компани. )|(?<=корпораци. )|(?<=фирм. ))\w+(.?бизнес|.?дизайн|.?авто|.?логистик|.?лизинг|.?нефтегаз|.?альянс|.?финанс)?"
mname_pattern = r"(?i)меня.*зовут|зовут.*меня|мое имя|да это|менеджер"
goodbye_pattern = r"(?i)до свидания|всего хорошего|всего доброго|хорошего дня|хорошего вечера"
hello_pattern = r"(?i)здравствуйте|добрый день|день добрый|добрый вечер|вечер добрый|доброе утро|приветствую"

df_new["say_goodbye"] = df_new.text.str.contains(goodbye_pattern)
df_new["my_name_is"] = df_new.text.str.contains(mname_pattern)
df_new["say_hello"] = df_new.text.str.contains(hello_pattern)
df_new["company"] = df_new.text.str.contains("компани|корпораци|фирм")


# СБОР ИМЁН МЕНЕДЖЕРОВ И НАЗВАНИЙ КОМПАНИЙ

regex_com = re.compile(company_pattern)
regex_man = re.compile(mname_pattern)
managers_names = []
companies_names = []

for text in df_new.text:
    if regex_com.search(text)!=None:
        company = regex_com.search(text).group(0)
        if company not in companies_names: companies_names.append(company)
    for word in text.split():
        for p in morph.parse(word):
            if ('Name' in p.tag) and (p.score>0.5) and ('sing' in p.tag) and ('nomn' in p.tag) and (
                regex_man.search(text)!=None):
                if word not in managers_names: managers_names.append(word)

                    
# ПОДГОТОВКА ОТЧЁТА

df_mini_report = pd.DataFrame(columns=["ID диалога", "приветствие", "прощание", "менеджер назвал имя", 
                                       "фраза с именем", "фраза с приветствием", 
                                       "фраза с прощанием", "название компании"])

for dlg_id in df_new.dlg_id.unique():
    dfb = df_new[df_new.dlg_id==dlg_id].reset_index()
    say_hello = False
    say_goodbye = False
    call_name = " - "
    my_name_phrase = " - "
    hello_phrase = " - "
    goodbye_phrase = " - "
    company_name = " - "
    for i in range(dfb.shape[0]):
        if dfb.iloc[i].say_hello:
            say_hello = True
            hello_phrase = dfb.iloc[i].text
        if dfb.iloc[i].say_goodbye:
            say_goodbye = True
            goodbye_phrase = dfb.iloc[i].text
        if dfb.iloc[i].my_name_is:
            for name in managers_names:
                if name in dfb.iloc[i].text:
                    call_name = name.capitalize()
            my_name_phrase = dfb.iloc[i].text
        if dfb.iloc[i].company:
            for one_company in companies_names:
                if one_company in dfb.iloc[i].text:
                    company_name = one_company
    df_mini_report.loc[dlg_id] = [dlg_id, say_hello, say_goodbye, call_name, my_name_phrase, 
                                  hello_phrase, goodbye_phrase, company_name]


# ВЫВОД ОТЧЁТА В ФАЙЛ 
 
df_mini_report.to_excel(r"mini_report.xlsx", index=False)
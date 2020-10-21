#!/usr/bin/env python
# coding: utf-8

# In[15]:


import os
import pandas as pd
from datetime import datetime as dt


# In[2]:


print(f'Текущая рабочая директория {(os.getcwd())}')


# In[3]:


#os.chdir('C:/Users/tabakaev_mv/Desktop/Сведение таблиц в одну')


# In[4]:


#print(f'Текущая рабочая директория {(os.getcwd())}')


# In[57]:


def excel_to_data_frame_parser(file):
    
    """
    Парсит эксель-файл.
    Возвращает датафрейм (далее - ДФ).
    """
    

    data = pd.ExcelFile(file)
    my_sheet_names = data.sheet_names
    list_number = 0
    df = data.parse(my_sheet_names[list_number])
    
    print(f'В ДФ с названием {file} {df.shape[0]} строк')
    print(f'В ДФ с названием {file} {df.shape[1]} столбцов')
    
    #print('OK!')
    return df


# In[69]:


def excel_concatenator(file_list, num_columns_to_remove: int = 8, num_to_drop_na: int = 12):
    
        
    for current_file in file_list:
        try:
            df = pd.DataFrame(columns=list(excel_to_data_frame_parser(current_file)))
            break
        except BaseException:
            print(f'Ошибка чтения таблицы из файла {current_file}')
            continue
    
        
    for file_name in file_list:
        
        if 'свод' in file_name.lower():
            continue
        
        try:
            
            current_df = excel_to_data_frame_parser(file_name)
            
        except BaseException:
            print(f'Ошибка чтения таблицы из файла {file_name}')
            continue
        
                
        names_identical = list(df) == list(current_df)
        
        if not(names_identical):
            print(f'В данном файле {file_name} названия столбцов не соответствуют таковым в первом файле')
        
        df = df.append(current_df)
        
        print(df.shape)
    
    
    print(f'Размерность ДФ ДО удаления строк {df.shape}')
    
    df = df.iloc[:,:-num_columns_to_remove]
    
    df = df.dropna(thresh=num_to_drop_na)
    
    print(f'Размерность ДФ ПОСЛЕ удаления строк {df.shape}')
    
    today_date = dt.today().strftime('%d.%m.%Y')
    
    print(f"Сегодняшняя дата: {today_date}")
    
    df.to_excel('СВОД_'+str(today_date)+'.xlsx', startrow=0, index=True)
    
    print(f'Файл эксель с названием "СВОД_{today_date}.xlsx" сохранен в текущей директории!')
    
    print('ВСЕГО ХОРОШЕГО!')
    input()
    
    return df


# In[63]:


file_list = os.listdir()
print(f'Список файлов для сведения: {file_list}')


# In[71]:


answer_go_forth = input('y/n')

num_col_del = input('Количество колонок справа к удалению (по умолчанию задано 8): ')

num_remove_in_row = input('Количество столбцов с пропущенными значениями для удаления всей строки (по умолчанию задано 12): ')



if 'n' not in answer_go_forth.lower() and (len(num_col_del)>0 and len(num_remove_in_row)>0):
    df = excel_concatenator(file_list, int(num_col_del), int(num_remove_in_row))
    
elif 'n' not in answer_go_forth.lower():
    df = excel_concatenator(file_list)
    
else:
    print('ВСЕГО ХОРОШЕГО!')
    input()


# In[ ]:





# In[ ]:





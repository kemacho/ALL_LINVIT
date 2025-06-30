
import os
from tqdm import tqdm
import shutil


# ПУТЬ К КОРНЕВОЙ ПАПКЕ
INpath = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2021'

# Названия папок
folder_names = os.listdir(INpath)


for i in tqdm(range(0, len(folder_names))):

    name = folder_names[i]
    pathSI = rf'{INpath}\{name}\0. СИ'
    pathIK1 = rf'{INpath}\{name}\1. ИК-1'
    pathIK2 = rf'{INpath}\{name}\2. ИК-2'


    path0 = rf'{pathSI}\0 Заявка и приложение'
    path1 = rf'{pathSI}\1 Распоряжение по заявке'
    path2 = rf'{pathSI}\2 Решение по заявке'
    path3 = rf'{pathSI}\3 Заключения по ОМД и ТД'
    path4 = rf'{pathSI}\4 Акт выбора ПК'
    path5 = rf'{pathSI}\5 Протоколы СИ'
    path6 = rf'{pathSI}\6 Заключение СИ'
    path7 = rf'{pathSI}\7 Программа проверки произ'
    path8 = rf'{pathSI}\8 Акт ПП'
    path9 = rf'{pathSI}\9 Распоряжение на анализ'
    path10 = rf'{pathSI}\10 Решение о выдаче'
    path11 = rf'{pathSI}\11 Сертификат'
    path12 = rf'{pathSI}\12 Доп.материалы'

    path3_1 = rf'{pathSI}\3 Заключения по ОМД и ТД\3.1 Заключение ОМД'
    path3_2 = rf'{pathSI}\3 Заключения по ОМД и ТД\3.2 Заключение ТД и РПН'
    path3_3 = rf'{pathSI}\3 Заключения по ОМД и ТД\3.3 Заключение ПМ'

    pathIK1_0 = rf'{pathIK1}\0 Распоряжение'
    pathIK1_1 = rf'{pathIK1}\1 Письмо-уведомление'
    pathIK1_2 = rf'{pathIK1}\2 Программа ИК'
    pathIK1_3 = rf'{pathIK1}\3 Программа проверки произ'
    pathIK1_4 = rf'{pathIK1}\4 Акт выбора ПК'
    pathIK1_5 = rf'{pathIK1}\5 Акт проверки производства'
    pathIK1_6 = rf'{pathIK1}\6 Протоколы ИК'
    pathIK1_7 = rf'{pathIK1}\7 Акт по результатам ИК'
    pathIK1_8 = rf'{pathIK1}\8 Распоряжение на анализ'
    pathIK1_9 = rf'{pathIK1}\9 Решение по ИК'
    pathIK1_10 = rf'{pathIK1}\10 Доп. материалы'

    pathIK2_0 = rf'{pathIK2}\0 Распоряжение'
    pathIK2_1 = rf'{pathIK2}\1 Письмо-уведомление'
    pathIK2_2 = rf'{pathIK2}\2 Программа ИК'
    pathIK2_3 = rf'{pathIK2}\3 Программа проверки произ'
    pathIK2_4 = rf'{pathIK2}\4 Акт выбора ПК'
    pathIK2_5 = rf'{pathIK2}\5 Акт проверки производства'
    pathIK2_6 = rf'{pathIK2}\6 Протоколы ИК'
    pathIK2_7 = rf'{pathIK2}\7 Акт по результатам ИК'
    pathIK2_8 = rf'{pathIK2}\8 Распоряжение на анализ'
    pathIK2_9 = rf'{pathIK2}\9 Решение по ИК'
    pathIK2_10 = rf'{pathIK2}\10 Доп. материалы'


    pathListSI = [path0, path1, path2, path3, path4, path5, path6, path7, path8, path9, path10, path11, path12]
    pathList_3 = [path3_1, path3_2, path3_3]
    pathlistIK1 = [pathIK1_0, pathIK1_1, pathIK1_2, pathIK1_3, pathIK1_4, pathIK1_5, pathIK1_6, pathIK1_7, pathIK1_8, pathIK1_9, pathIK1_10]
    pathlistIK2 = [pathIK2_0, pathIK2_1, pathIK2_2, pathIK2_3, pathIK2_4, pathIK2_5, pathIK2_6, pathIK2_7, pathIK2_8, pathIK2_9, pathIK2_10]


    #Создание папки СИ и ее содержимого
    if not os.path.exists(pathSI):
        for j in range(0, len(pathListSI)):
            if j == 3 and not os.path.exists(pathListSI[j]):
                os.makedirs(pathListSI[j])
                os.makedirs(pathList_3[0])
                os.makedirs(pathList_3[1])
                os.makedirs(pathList_3[2])
            elif not os.path.exists(pathListSI[j]):
                os.makedirs(pathListSI[j])

    # Создание папки ИК1 и ее содержимого
    for j in range(0, len(pathlistIK1)):
        if not os.path.exists(pathlistIK1[j]):
            os.makedirs(pathlistIK1[j])

    # Создание папки ИК2 и ее содержимого
    for j in range(0, len(pathlistIK2)):
        if not os.path.exists(pathlistIK2[j]):
            os.makedirs(pathlistIK2[j])

        
print('Successful!')
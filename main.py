import pandas as pd
import pathlib
from pathlib import Path
import random
from graphviz import Digraph
from openpyxl import load_workbook
import os
import warnings
warnings.filterwarnings('ignore')

# отключаем лимит на строки
pd.set_option('display.max_rows', None)
# отключаем лимит на колонки
pd.set_option('display.max_columns', None)
# отключаем лимит на количество символов в записи
pd.set_option('display.max_colwidth', None)
# отключаем перенос при выводе записи на экран(тут в принте), иначе перенесет на новую строку
pd.options.display.expand_frame_repr = False

# Получаем путь, можешь посмотреть work_path через принт
work_path = pathlib.Path.cwd()

# path1 = Path(work_path, 'Выгрузка СХ - копия.xlsx')
path2 = Path(work_path, 'Архивные.xlsx')
path3 = Path(work_path, 'Исходные.xlsx')

# Читаем файлы, v  - выгрузка СХ(нужна была для функции чистки)
# v = pd.read_excel(path1)
arhiv = pd.read_excel(path2)
ishod_df = pd.read_excel(path3)

# Разкоментить, если нужна функция чистки
# Сравнивает  кн в архивном файле и выгрузке и удаляет строки из выгрузки, при совпадении
# def chistka(v,arhiv):
#     v = v
#     arhiv = arhiv
#     for i in v["Исходный номер исх. ЗУ"]:
#         if i in arhiv['КН ЗУ'].unique():
#             spisok_indeksov = v.index[v['Исходный номер исх. ЗУ'] == i].tolist()
#             if len(spisok_indeksov) > 1:
#                 for ind in spisok_indeksov:
#                     v.drop(labels=[ind], axis=0, inplace=True)
#                     print(f'Удален участок цикл {i}')
#             elif len(spisok_indeksov) == 1:
#                 v.drop(labels=[*spisok_indeksov], axis=0, inplace=True)
#                 print(f'Удален участок {i}')
# # chistka(v,arhiv)
# # v.to_excel(path1, index= False) Если надо, записываем в тот же эксель, что и читали

# # Ищем действительно исходные и записываем в отдельный эксель файл
# sp_chist = []
# uniq_ishod = [isxod_zy for isxod_zy in arhiv["КН исходного"].unique()]
# uniq_posl = [posl for posl in arhiv["КН ЗУ"].unique()]
# for isxod_zy in uniq_ishod[1:]:
#     if isxod_zy in uniq_posl[1:]:
#         continue
#     else:
#         sp_chist.append(isxod_zy)
#         print(f"Зу {isxod_zy} добавлен в список")
# ishod_df["Исходные Зу"] = pd.Series(sp_chist)
# ishod_df.to_excel(path3, index= False)


df_sheets = {} # Словарь для добавления разных листов зу, в один эксель
# Cписок уникальных земельный участков
zem_uch = [i for i in ishod_df["Исходные Зу"]]
kvartal = {str(i.split('_')[0] +'_'+ i.split('_')[1] + '_'+ i.split('_')[2]) for i in zem_uch if isinstance(i,str)}
# Проходимся по множеству кварталов и добавляем в список участки с однаковым кварталом
for kv in kvartal:
    if not os.path.isdir(f'.//Готовые_деревья/{kv}.xlsx'):
        try:
            os.mkdir(f".//Готовые_деревья/{kv}")
        except FileExistsError:
            continue
        sp = [zem for zem in zem_uch if str(zem.split('_')[0] +'_'+ zem.split('_')[1] + '_'+ zem.split('_')[2]) == kv]

        #  Добавляем исходный участок, помещаем в new_df и ищем его детей
        for i in sp:
            print(f"Начало работы с кварталом {kv}, зу {sp[sp.index(f'{i}')]}")
            new_df = pd.DataFrame()
            # Зу для проверки - 50_13_0030417_36   50_13_0000000_252  50:18:0000000:113  50_16_0203013_8 50_16_0000000_55 50_16_0000000_35 50_06_0000000_65


            #  Получаем индексы исходного зу
            index_isxod_arhiv = arhiv.index[arhiv['КН исходного'] == i].tolist()

            # Создаем столбцы в дф, добавляем значения индексов для исходного и ребенка
            new_df["Исходный зу"] = [arhiv['КН исходного'].values[i] for i in index_isxod_arhiv]
            new_df["Площадь"] = arhiv['Площадь исходного'].values[index_isxod_arhiv[0]]
            new_df["Право"] = arhiv['статус исходного'].values[index_isxod_arhiv[0]]


            count_zy = 0

            new_df[f"Последующий Зу_{count_zy}"] = [arhiv['КН ЗУ'].values[i] for i in index_isxod_arhiv]
            new_df[f"Площадь Зу_{count_zy}"] = [arhiv['Площадь'].values[i] for i in index_isxod_arhiv]
            new_df[f"Право Зу_{count_zy}"] = [arhiv['статус'].values[i] for i in index_isxod_arhiv]
            # for kolonka in new_df:
            #     new_df[kolonka] = pd.Series(new_df[kolonka].unique())
            # new_df = new_df.drop_duplicates()

            # Удаляем дубликаты и сбрасываем индексы
            new_df = new_df.drop_duplicates()
            new_df.reset_index(drop=True, inplace=True)







            # # Получаем индексы ребенка
            # index_detei_ishod = [arhiv.index[arhiv['КН исходного'] == i].tolist() for i in new_df[f"Последующий Зу_{count_zy}"]]
            #
            #
            #
            # # Создаем временные серии для будущей колонки внука (зу, площадь и право)
            # temp_zy = pd.Series()
            # temp_s = pd.Series()
            # temp_pravo = pd.Series()
            # kolvo = 0
            # # Добавляем значения индексов ребенка и значения индекса внука в одну серию
            # for detei_ishod in index_detei_ishod:
            #     temp_zy = temp_zy.append(pd.Series([[arhiv['КН ЗУ'].values[i] for i in detei_ishod]]))
            #     # Сбрасываем индексы
            #     temp_zy.reset_index(drop=True, inplace=True)
            #     for qq in index_isxod_arhiv[kolvo]:
            #         temp_zy[kolvo].append(arhiv['КН образованного'].values[qq])
            #
            #     temp_s = temp_s.append(pd.Series(arhiv['Площадь'].values[i] for i in detei_ishod))
            #     temp_s = temp_s.append(pd.Series(arhiv['Площадь образованного'].values[index_isxod_arhiv[kolvo]]))
            #     temp_s.reset_index(drop=True, inplace=True)
            #     temp_pravo = temp_pravo.append(pd.Series([arhiv['статус'].values[i] for i in detei_ishod]))
            #     temp_pravo = temp_pravo.append(pd.Series(arhiv['статус образованного'].values[index_isxod_arhiv[kolvo]]))
            #     temp_pravo.reset_index(drop=True, inplace=True)
            #     kolvo += 1
            # kolvo = 0
            #
            # # Создаем новые столбцы и добавляем в них серии
            # count_zy += 1
            # new_df[f"Последующий Зу_{count_zy}"] = temp_zy
            # new_df = new_df.explode(f"Последующий Зу_{count_zy}", ignore_index=True)
            # new_df[f"Площадь Зу_{count_zy}"] = temp_s
            # new_df[f"Право Зу_{count_zy}"] = temp_pravo
            #
            # # Удаляем дубликаты и сбрасываем индексы
            # new_df = new_df.drop_duplicates()
            # new_df.reset_index(drop=True, inplace=True)

            while True:
                # Получаем индексы внука
                index_vnuk_ishod = [arhiv.index[arhiv['КН исходного'] == i].tolist() for i in new_df[f"Последующий Зу_{count_zy}"]]
                index_obrazovan= [arhiv.index[arhiv['КН ЗУ'] == i].tolist() for i in new_df[f"Последующий Зу_{count_zy}"]]
                index_obrazovan = [[0] if i == [] else i for i in index_obrazovan]
                proverka_ind = [x for i in index_obrazovan for x in i]
                if sum(proverka_ind) == 0:
                    break
                else:
                    # Создаем временные серии для будущей колонки внука (зу, площадь и право)
                    t_zy = pd.Series()
                    t_s = pd.Series()
                    t_pravo = pd.Series()
                    kolvo = 0
                    for vnul_ishod in index_vnuk_ishod:
                        t_zy = t_zy.append(pd.Series([[arhiv['КН ЗУ'].values[i] for i in vnul_ishod]]))
                        t_zy.reset_index(drop=True, inplace=True)
                        for m in index_obrazovan[kolvo]:
                            if m != 0:
                                t_zy[kolvo].append(arhiv['КН образованного'].values[m])

                        if m == 0 and len(vnul_ishod) == 0:
                            t_s = t_s.append(pd.Series([None]))
                        else:
                            t_s = t_s.append(pd.Series([arhiv['Площадь'].values[i] for i in vnul_ishod]))
                            if sum(index_obrazovan[kolvo]) != 0:
                                t_s = t_s.append(pd.Series(arhiv['Площадь образованного'].values[index_obrazovan[kolvo]]))
                        t_s.reset_index(drop=True, inplace=True)
                        if m == 0 and len(vnul_ishod) == 0:
                            t_pravo = t_pravo.append(pd.Series([None]))
                        else:
                            t_pravo = t_pravo.append(pd.Series([arhiv['статус'].values[i] for i in vnul_ishod]))
                            if sum(index_obrazovan[kolvo]) != 0:
                                t_pravo = t_pravo.append(pd.Series(arhiv['статус образованного'].values[index_obrazovan[kolvo]]))
                        t_pravo.reset_index(drop=True, inplace=True)
                        kolvo += 1
                    kolvo = 0

                    # Создаем новые столбцы и добавляем в них серии
                    count_zy += 1
                    new_df[f"Последующий Зу_{count_zy}"] = t_zy
                    new_df = new_df.explode(f"Последующий Зу_{count_zy}", ignore_index=True)
                    new_df[f"Площадь Зу_{count_zy}"] = t_s
                    new_df[f"Право Зу_{count_zy}"] = t_pravo


                    # Удаляем дубликаты и сбрасываем индексы
                    new_df = new_df.drop_duplicates()
                    new_df.reset_index(drop=True, inplace=True)

            # # x = x.fillna("ychastok_otsutstvuet") # Разкоментить, если надо заменить значение nan
            # # # print(x.head(30)) # Показывает первые 30 строк датафрейма.

            #Добавляем датафреймы в словарь, ключ - кадастровый номер(он же будет названием листа)
            df_sheets[f'{i}'] = new_df

        print(f"Окончание работы со списком, по кварталу {kv}, длина словаря {len(df_sheets)}, идет запись в листы")

        # Записываем фреймы в разные листы и сохраняем
        # x.to_excel(r'C:\Users\denis.osipov\PycharmProjects\DEREVO\zy.xlsx', index= False)
        writer = pd.ExcelWriter(fr'C:\Users\denis.osipov\PycharmProjects\DEREVO\Готовые_деревья\{kv}\{kv}.xlsx', engine='xlsxwriter')
        for sheet_name in df_sheets.keys():
            df_sheets[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)
            for column in df_sheets[sheet_name]:
                column_width = max(df_sheets[sheet_name][column].astype(str).map(len).max(), len(column))
                col_idx = df_sheets[sheet_name].columns.get_loc(column)
                writer.sheets[sheet_name].set_column(col_idx, col_idx, column_width)
        df_sheets = dict()
        writer.close()
        print(f"Окончание работы с листами, словарь очищен, квартал {kv} записан, идет создание картинки")


    # # Строим логику для графики
    # """Пробелов быть не должно между словами. Все слова на англ. языке. ":" не принимаются, ставь "_" """
    #
    # # считываем файл
    # df = pd.read_excel(fr'C:\Users\denis.osipov\PycharmProjects\DEREVO\Готовые_деревья\{kv}\{kv}.xlsx', dtype=str, sheet_name=sp)
    # count_listov = 0
    # for sheet in df:
    #     sheet_df = df[sheet]
    #     # разносим таблицы и столбцы в два dataframe
    #     df1 = sheet_df[["Исходный зу", "Площадь", "Право"]]
    #     df_posl = sheet_df[[column for column in sheet_df][3:]]
    #     # Узнаем длину последующих участков
    #     num = int(len(df_posl.columns)/3)
    #     spisok_df = []
    #     start_df = 0
    #     end_df = 3
    #     # Создаем для каждого последующего участка датафрейм и заносим всех в список
    #     for count in range(num):
    #         df2 = sheet_df[[column for column in df_posl][start_df:end_df]]
    #         spisok_df.append(df2)
    #         start_df = start_df + 3
    #         end_df = end_df + 3
    #     # для объединения переименуем столбцы
    #     # Убираем хвосты, получаем ощиченный список последующих уч.
    #     number_df = 0
    #     chist_sp_df = []
    #     for frame in spisok_df:
    #         x = frame.columns.values[0][-2:]
    #         frame.columns = frame.columns.str.replace(f'{x}', '')
    #         frame = frame.rename(columns={'Последующий Зу': 'Исходный зу', 'Площадь Зу': 'Площадь', 'Право Зу': 'Право'})
    #         chist_sp_df.append(frame)
    #
    #
    #     # объединяем  dataframe в один и убираем дубликаты
    #     df_concat = pd.concat([df1, *chist_sp_df])
    #     df_concat = df_concat.drop_duplicates() # Удаляются дубликаты
    #     df_concat = df_concat.sort_values(by='Исходный зу') # Сортировка по столбцу Т1, по возрастанию
    #     df_concat  = df_concat.reset_index(drop=True) # Сбрасывается индекс
    #     # для создания node
    #     df_concat['Исходный зу_new'] = '[' + df_concat['Исходный зу'] + ']' # Создается дополнительный столбец Т1_new со знач из Т1
    #     df_concat['Площадь_new'] = '|<' + df_concat['Площадь'] + '> ' + df_concat['Площадь'] # Создается дополнительный столбец S1_new со знач из S1
    #     df_concat['Право_new'] = '|<' + df_concat['Право'] + '> ' + df_concat['Право']
    #     df_concat = df_concat.drop(['Исходный зу', 'Площадь', 'Право'], axis=1) # Удаляются оригинальные столбцы Т1, S1
    #     # Создаем два новых столбца для Площади и Права
    #     df_concat['Count'] = df_concat.groupby('Исходный зу_new').cumcount()
    #     df_concat['Count_1'] = df_concat.groupby('Исходный зу_new').cumcount()
    #     # Исходный_new стал индексом, Count, Count_1 - столбцами , Площадь_new и Право_new - строками этих столбцов соответственно
    #     df_pivot = df_concat.pivot('Исходный зу_new', ['Count', 'Count_1'], ['Площадь_new','Право_new'])
    #     # Преобразование DataFrame в массив записей NumPy
    #     df_p = pd.DataFrame(df_pivot.to_records())
    #     # Создается нов. столбец concat, где Исходный_new + значения из столбц. "0" помещаются в одну строку
    #     df_p['concat'] = pd.Series(df_p.fillna('').values.tolist()).str.join('')
    #     # Удаляется столбец "0", остаются Исходный_new и concat
    #     df_node_label = df_p[['Исходный зу_new', 'concat']]
    #     # Добавляется столбец Res , где в его значения помещаются данные из Исходный_new и concat, а между ними знак $$$
    #     df_node_label['Res'] = pd.Series(df_node_label.fillna("").values.tolist()).str.join("$$$")
    #     df_node_label = df_node_label.dropna()
    #     df_node_label  = df_node_label.reset_index(drop=True)
    #     # для связей
    #     # Добавляется столбец Res в изн. фрейм дф, указываем связи Исходный:площадь:право$$$Последующий:площадь:право и в цикле все участки соединияются к исходному
    #     sheet_df['Res'] = '[' + sheet_df['Исходный зу'] + ']' + ':' + sheet_df['Площадь'] + ':' + sheet_df['Право']
    #     for l in range(num):
    #         sheet_df['Res'] = sheet_df['Res'] +'$$$' +'[' + sheet_df[f'Последующий Зу_{l}'].fillna('') + ']' + ':' + sheet_df[f'Площадь Зу_{l}'].fillna('') + ':' + sheet_df[f'Право Зу_{l}'].fillna('')
    #
    #     ar1 = df_node_label['Res'].values
    #     ar2 = sheet_df['Res'].values
    #
    #     # отрисовка
    #     g = Digraph('structs') # format='jpg' или format='png'. По-умолчанию - pdf
    #     # g.attr(rankdir="LR") # Если включить, стрелки будут слева на право
    #     g.attr('node', shape= "box") #  shape='plaintext' - форма нода
    #     g.attr(size='10000')
    #
    #
    #     # рандомный цвет для ребер
    #     def rcolor():
    #         r = random.randint(0, 16777215)
    #         hexnumber = str(hex(r))
    #         hexnumber = '#' + hexnumber[2:]
    #         return hexnumber
    #
    #     # Создаем ноды
    #     for i in ar1:
    #         node_ = i.split("$$$")[0]
    #         node_label = i.split("$$$")[1]
    #         g.node(node_, label=node_label)
    #
    #
    #     # Создаем связи между нодами
    #     num_nod = 0
    #     for sp_edges in ar2:
    #         yzel = sp_edges.split("$$$")
    #         for dot in range(len(yzel)-1):
    #             edge1 = yzel[num_nod]
    #             edge2 = yzel[num_nod + 1]
    #             if edge1 == '[]::' or edge2 == '[]::':
    #                 continue
    #             else:
    #                 # можно менять параметр minlen для изменения расстояния между nodes
    #                 g.edge(edge1, edge2, color = rcolor(), minlen='5')
    #                 num_nod += 1
    #         num_nod = 0
    #
    #     # Cохранить схему, view=True - покажет ее
    #     g.render(directory=f".//Готовые_деревья/{kv}/{sheet}", view=False, filename=f'{kv}')
    #     count_listov += 1
    # count_listov = 0
    #     # filename = g.render(filename=f'{kv}')
    #     # #Показать и сохранить схему
    #     # g.view()
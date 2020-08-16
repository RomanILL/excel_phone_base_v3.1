import os
from glob import glob
import func_xlsx_base
from make_phone_func import make_good_phone_list
import find_address_base

if __name__ == "__main__":

    # блок констант
    version = 3.1

    global_count = 0
    COUNTRY_SELECT = ["Россия"]
    exit_phone_base_file_name = "Томская-Кемеровская.xlsx"
    dir_input_files = "input_xlsx"
    dir_output_files = "output_xlsx"
    dir_supporting_files = "support_tables"
    support_cities_table_name = "cities.xlsx"
    support_regions_table_name = "regions.xlsx"
    itn_original_head_id = 6  # individual tax number - ИНН (номер колонки в оригинальном файле)
    number_vehicle_original_head_id = 19
    sign_list_original_head_id = (5, 15, 16, 20)
    sign_list_edit_head_id = (1, 7, 8, 11)
    ignore_cols_list_original_head_id = (1, 2, 11, 12, 13, 14, 18, 21, 22, 23, 24, 25, 27)

    filter_cities = ["Томская обл.", "Кемеровская обл.", "Томская область", "Кемеровская область", "ТО", "КО"]

    # индексы заголовков, где что лежит
    # FIXME переделать в удобный блок или словарь

    """
    parent_city_head_id = 0 родитель
    address_registration_head_id = 4 юр.адрес
    actual_address_head_id = 5 факт. адрес
    phones_list_head_id = 7 список телефонов
    vehicle_number_head_id = 11 номер ТС

    """

    parent_city_head_id = 0
    address_registration_head_id = 4
    actual_address_head_id = 5
    phones_list_head_id = 7
    vehicle_number_head_id = 11


    # проверяем папки назначения (откуда брать, куда класть, вспомогательная)
    func_xlsx_base.check_destination_folders((dir_input_files, dir_output_files, dir_supporting_files))

    # собираем список файлов для обработки
    phone_base_files_list = glob(f'{dir_input_files}\\*.xlsx')
    # печатаем список файлов для обработки - FIXME можно убрать после отладки
    func_xlsx_base.print_any_list(phone_base_files_list)
    print("-" * 30 + "\n")

    # формируем список заголовков таблиц из первого файла - перенесено в главный цикл программы

    # список - добавка в шапку таблиц
    # FIXME после отладки блока принятия решений оставить только город / регион, страну и телефон по формату
    extend_head_list = ["Кол-во ТС", "Страна", "Город Юр.лица", "Город Факт.", "Регион", "Регион ТС",
                        "Первый телефон", "Доп. номера "]

    # формируем вспомогательные данные (словари городов и номеров регионов)
    # формируем словарь городов
    cities_dict = func_xlsx_base.make_cities_dict(dir_supporting_files + "\\" + support_cities_table_name, COUNTRY_SELECT)
    # формируем словарь регионов
    region_dict = func_xlsx_base.make_regions_dict(dir_supporting_files + "\\" + support_regions_table_name)
    # формируем словарь названий регионов
    region_names_set = func_xlsx_base.make_regions_names_set(dir_supporting_files + "\\" + support_cities_table_name, COUNTRY_SELECT)

    # перебираем исходные файлы (признак первого открытого файла - пустая переменная heads)
    heads = None
    for current_file_name in phone_base_files_list:
        if "~" not in current_file_name:
            temp_file_name = current_file_name.split("\\")[-1]
            print(f'Открываем файл для чтения: {temp_file_name}')
            current_xlsx_obj = func_xlsx_base.open_file_xlsx(current_file_name)
            if heads is None:
                # если это первый файл из списка, то собираем заголовки

                print(f"Читаем заголовки таблицы {temp_file_name}")
                heads = func_xlsx_base.get_heads(current_xlsx_obj, ignore_id_list=ignore_cols_list_original_head_id)
                # печатаем список заголовков


                # открываем файл для записи
                if os.path.isfile(dir_output_files + "\\" + exit_phone_base_file_name):
                    # delete current file
                    print(f'Удаляется ранее созданный файл "{exit_phone_base_file_name}"')
                    os.remove(dir_output_files + "\\" + exit_phone_base_file_name)
                exit_phone_xlsx = func_xlsx_base.create_write_file(dir_output_files + "\\" + exit_phone_base_file_name)
                # расширяем заголовки
                heads.extend(extend_head_list)
                # прописываем заголовки в новый файл
                exit_phone_xlsx.active.append(heads)

            # блок обработки файла
            # берем первый по индексу лист
            current_xlsx_obj.active = 0
            active_sheet = current_xlsx_obj.active
            numbers_rows = active_sheet.max_row
            numbers_cols = active_sheet.max_column
            list_of_unique_itn = []


            # собираем строку для записи
            for i_row in range(2, numbers_rows + 1):

                # проверяем на уникальность ИНН
                current_check_itn =str(active_sheet.cell(row=i_row, column=itn_original_head_id).value)
                if current_check_itn in list_of_unique_itn:
                    continue
                else:
                    # блок добавляет немного магии в сбор данных
                    # кол-во машин у одного контрагента (при условии сортировки по контрагенту)
                    # найдем из длины списка машин

                    # список машин будем закидывать в ячейку "номер ТС"
                    vehicle_number_list = []
                    # сделаем собиратель статусов ( если хоть один "да" - общий статус "да") для аренды и т.п.
                    sign_list = [False, False, False, False]
                    temp_itn = current_individual_tax_number = str(
                        active_sheet.cell(row=i_row, column=itn_original_head_id).value)
                    temp_index_row = i_row
                    while temp_itn == current_individual_tax_number and current_individual_tax_number.strip() != "":
                        # собираем список машин каждого контрагента

                        vehicle_number_list.append(
                            str(active_sheet.cell(row=temp_index_row,
                                                  column=number_vehicle_original_head_id).value).strip())
                        # хитрый механизм сбора признаков - двумерная структура сначала проходит по строке,
                        # а затем добавляет элемент след строки, пока не станет True
                        for sign_id in range(len(sign_list)):
                            if sign_list[sign_id] is False:
                                if str(active_sheet.cell(row=temp_index_row,
                                                         column=sign_list_original_head_id[sign_id]).value)[
                                    0] == "Д":
                                    sign_list[sign_id] = True

                        temp_index_row += 1
                        temp_itn = str(active_sheet.cell(row=temp_index_row, column=itn_original_head_id).value)
                    # теперь у нас есть список машин и список признаков

                    list_of_unique_itn.append(current_check_itn)

                row_for_write = []
                extend_row = [None, None, None, None, None, None, None, None]
                """ "Кол-во ТС", "Страна", "Город Юр.лица", "Город Факт.", "Регион", "Регион ТС", "Первый телефон", "Доп. номера " """
                """ [27, 28, 29, 30, 31, 32, 33, 34] """
                for j_col in range(1, numbers_cols + 1):
                    # индексы заголовков, где что лежит
                    """
                    parent_city_head_id = 0 родитель
                    address_registration_head_id = 4 юр.адрес
                    actual_address_head_id = 5 факт. адрес
                    phones_list_head_id = 7 список телефонов
                    vehicle_number_head_id = 11 номер ТС

                    """


                    # собираем новую строку, исключая ненужные столбцы
                    if j_col not in ignore_cols_list_original_head_id:
                        current_cell_value = str(active_sheet.cell(row=i_row, column=j_col).value)

                        # список машин ставим в ячейку "номер ТС", когда цикл дойдет до нее
                        if j_col == number_vehicle_original_head_id:
                            current_cell_value = ", ".join(vehicle_number_list)

                        # ставим соответствующий признак, когда цикл доходит до признака
                        elif j_col in sign_list_original_head_id:
                            for sign_id in range(len(sign_list_original_head_id)):
                                if j_col == sign_list_original_head_id[sign_id]:
                                    if sign_list[sign_id]:
                                        current_cell_value = "Да"
                                    else:
                                        current_cell_value = "Нет"
                                break

                        row_for_write.append(current_cell_value)


                # в этом моменте у нас собралась строка для записи из оригинального файла
                # чистим строку от None

                for temp_value_id in range(len(row_for_write)):
                    if row_for_write[temp_value_id] is None:
                        row_for_write[temp_value_id] = ""


                # собираем строку extend_row из оригинала

                # extend_name_cell_dict = encd[]
                encd = {"Кол-во ТС": 0, "Страна": 1, "Город Юр.лица": 2, "Город Факт.": 3,
                                         "Регион": 4, "Регион ТС": 5, "Телефон": 6, "Доп. номера": 7}
                current_row_cities_for_filter = []

                extend_row[encd["Кол-во ТС"]] = len(vehicle_number_list)
                # ищем города в адресах
                for city_name in cities_dict:
                    if city_name in row_for_write[address_registration_head_id]:
                        extend_row[encd["Город Юр.лица"]] = city_name
                    if city_name in row_for_write[actual_address_head_id]:
                        extend_row[encd["Город Факт."]] = city_name
                # определяем регион и страну по городу (приоритет фактическому адресу)
                if extend_row[encd["Город Факт."]] is not None:
                    extend_row[encd["Страна"]] = cities_dict[extend_row[encd["Город Факт."]]][0]
                    extend_row[encd["Регион"]] = cities_dict[extend_row[encd["Город Факт."]]][1]
                    # добавка для фильтра
                    current_row_cities_for_filter.append(cities_dict[extend_row[encd["Город Факт."]]][1])
                elif extend_row[encd["Город Юр.лица"]] is not None:
                    extend_row[encd["Страна"]] = cities_dict[extend_row[encd["Город Юр.лица"]]][0]
                    extend_row[encd["Регион"]] = cities_dict[extend_row[encd["Город Юр.лица"]]][1]
                    # добавка для фильтра
                    current_row_cities_for_filter.append(cities_dict[extend_row[encd["Город Юр.лица"]]][1])

                # определяем регион, если город не нашелся в адресе
                if extend_row[encd["Регион"]] is None:
                    for region_name in region_names_set:
                        if region_name in row_for_write[actual_address_head_id]:
                            extend_row[encd["Регион"]] = region_name
                            break
                        if region_name in row_for_write[address_registration_head_id]:
                            extend_row[encd["Регион"]] = region_name
                            break

                # определяем регион по номеру машины
                temp_region_list =[]
                for tmp in vehicle_number_list:
                    temp_region_list.append(tmp.split('-')[-1])
                # print(temp_region_list)
                region = func_xlsx_base.most_frequent(temp_region_list)
                # print(region)

                if str(region) in region_dict:
                    extend_row[encd["Регион ТС"]] = region_dict[region]
                # print(extend_row[encd["Регион ТС"]])
                # добавка для фильтра
                current_row_cities_for_filter.append(region_dict[region] if region in region_dict else "")

                # блок работы с телефонами
                # находим сотовые телефоны и приводим их к стандартному виду
                phone_numbers_string = row_for_write[phones_list_head_id]
                # получаем список мобильных телефонов компании
                company_phone_list = make_good_phone_list(phone_numbers_string)

                row_for_write.extend(extend_row)

                # делаем ячейки с телефонами
                # print(heads)
                # print(len(row_for_write))
                if len(company_phone_list) > 0:
                    row_for_write[20] = company_phone_list[0]
                    if len(company_phone_list) > 1:
                        row_for_write[21] = ", ".join(company_phone_list[1:])

                    # чистим от None
                    for cell_id in range(len(row_for_write)):
                        if row_for_write[cell_id] is None:
                            row_for_write[cell_id] = ""

                    # фильтр на город - пишем в файл только те, которые соответствуют обозначенной области,
                    # если нужно все, то оставляем только строку с append

                    for city_n in current_row_cities_for_filter:
                        if city_n in filter_cities:
                            if global_count % 50 == 0:
                                print(f"кол-во записей - {global_count}.")
                            exit_phone_xlsx.active.append(row_for_write)
                            global_count += 1
                            if global_count > 800000:
                                print("кол-во записей превысило 800 тыс.")
                            break

                """если нужно разделить каждый номер на отдельную строку - использовалось раньше - сейчас не нужно
                for write_phone in company_phone_list:
                    row_for_write[20] = write_phone
                    exit_phone_xlsx.active.append(row_for_write)
                """

            # блок закрытия файла после использования
            current_xlsx_obj.close()
        else:
            pass
    exit_phone_xlsx.save(dir_output_files + "\\" + exit_phone_base_file_name)
    print(f"В файле {exit_phone_base_file_name} сохранены изменения")
    exit_phone_xlsx.close()
    print("Программа успешно завершена.")
    input('Для заверения нажмите Enter...')

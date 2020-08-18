from glob import glob
import config_data as co_da
import find_address_base as fab
import other_function_xlsx
import os
import phone_func

if __name__ == "__main__":
    print("Запущена программа сбора базы из нескольких баз.")
    # блок основных настроек прописан в config_data (cd.)
    # создадим счетчик обработанных записей (всего)

    counter_processed_rows = 0
    print("Собираем список файлов для обработки")
    # собираем список файлов для обработки
    phone_base_files_list = glob(f'{co_da.dir_input_files}\\*.xlsx')
    print("Файлов найдено:", len(phone_base_files_list))
    # печатаем список файлов для обработки - убрано после отладки
    # other_function_xlsx.print_any_list(phone_base_files_list)
    # print("-" * 30 + "\n")
    print("Инициализация программных ресурсов")
    print("-" * 30)

    # раньше мы формировали заголовки из оригинального файла, а теперь будем задавать жестко, т.к.
    # требования к столбцам совершенно определенные(для скорости используем кортеж)
    heads_names = "Страна", "Фед. округ", "Регион/обл.", "Контрагент", "ИНН", "Вид деятельности", "Кол-во ТС", \
                  "Юр. адрес", "Город юр.адр.", "Факт. адрес", "Город ф.адр.", "Список ТС", "Регион ТС", \
                  "Телефон для СМС", "Доп. моб. номера", "Список номеров в оригинале", "Аренда", \
                  "Свои ТС", "Не выводить", "Не звонить", "Комментарий", "Рассписание работы", \
                  "Нужна ручная проверка адреса"

    # (сейчас просто пишет в отдельный лист)
    # список игнорируемых столбцов из оригинального файла (вероятно не потребуется, так как оформлять таблицу будем
    # с жёсткой привязкой и по этапам
    ignore_cols_list_original_head_id = (1, 2, 3, 11, 12, 13, 14, 18, 21, 22, 23, 24, 25, 5, 15, 16, 19, 20)
    # создадим вспомогательные списки для работы с адресами
    support_address_object = fab.RegionDistrict()
    """
    print("собранный словарь регионов ТС")
    for i in fab.RegionDistrict.vehicle_code_dict:
        print('{} - {}'.format(i, fab.RegionDistrict.vehicle_code_dict[i]))
        """
    # Теперь у нас все переменные созданы и инициализированы в классе и можно обращаться к методам класса
    # (FIXME объект нам не понадобится - пожалуй, стоит сделать рефакторинг без классов через обычные функции)

    # федеральных округов для удобства доступен через класс fab.RegionDistrict.district_tuple
    # нам он понадобится, когда будем переключаться между вкладками

    # открываем файл для записи
    if os.path.isfile(co_da.dir_output_files + "\\" + co_da.exit_phone_base_file_name):
        # delete current file
        print(f'Удаляется ранее созданный файл "{co_da.exit_phone_base_file_name}"')
        os.remove(co_da.dir_output_files + "\\" + co_da.exit_phone_base_file_name)
    exit_phone_xlsx = other_function_xlsx.create_write_file(co_da.dir_output_files + "\\" +
                                                            co_da.exit_phone_base_file_name)
    print("-" * 30)

    # сюда добавим блок создания вкладок по количеству федеральных округов (с заголовками)
    for current_district_name_id in range(len(fab.RegionDistrict.district_tuple) - 1):
        exit_phone_xlsx.active = current_district_name_id
        exit_phone_xlsx.active.append(heads_names)
        # если первый лист, то он будет переименован (чтобы не удалять и создавать заново)
        if current_district_name_id == 0:
            exit_phone_xlsx.active.title = fab.RegionDistrict.district_tuple[0]

        exit_phone_xlsx.create_sheet(fab.RegionDistrict.district_tuple[current_district_name_id + 1])
    # переключаемся и записываем заголовки в крайний созданный лист
    exit_phone_xlsx.active = len(fab.RegionDistrict.district_tuple) - 1
    exit_phone_xlsx.active.append(heads_names)

    # exit_phone_xlsx.save(co_da.dir_output_files + "\\" + co_da.exit_phone_base_file_name)

    # после не забываем сохранить и закрыть файл - совсем в конце, пока, вероятно, можно убрать
    # exit_phone_xlsx.close()

    # заведем словарь для именования ячеек
    # origin_cell_num_dict = oc_num номера ячеек в оригинальном файле

    oc_num = {"Контрагент": 4, "Аренда": (5, 16), "ИНН": 6, "Юр. адрес": 7, "Факт. адрес": 8,
              "Вид деятельности": 9,
              "Телефон": 10, "Не выводить": (15, 18), "Не звонить": (16, 19), "Комментарий": 17,
              "Номер ТС": (19, 11),
              "Своя ТС": (20, 17), "Рассписание": 26}
    new_num = {"Кол-во ТС": 6, "Регион ТС": 12, "СМС": 13, "Доп. номера": 14,
               "Страна": 0, "Округ": 1, "Регион": 2, "Юр. город": 8, "Факт. город": 10,
               "Адрес не распознан": 22}

    just_transfer_list = {4: 3, 6: 4, 7: 7, 8: 9, 9: 5, 10: 15, 17: 20, 26: 21}
    # эти данные можно смело вписывать в игнор лист и собирать по ним данные в блоке сбора данных отдельно FIXME
    data_processing_list = (5, 7, 8, 10, 15, 16, 19, 20)

    sign_list_names = ("Аренда", "Своя ТС", "Не выводить", "Не звонить")
    # put cell num dict = pc_num номера ячеек в собираемом файле

    # начинаем перебирать файлы из папки исходных файлов
    for current_file_name in phone_base_files_list:
        # будем исключать временные файлы
        if "~" not in current_file_name:

            temp_file_name = current_file_name.split("\\")[-1]
            counter_from_file = 0
            print(f'Обрабатывается файл: {temp_file_name}')
            current_xlsx_obj = other_function_xlsx.open_file_xlsx(current_file_name)

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
                current_check_itn = str(active_sheet.cell(row=i_row, column=oc_num["ИНН"]).value)
                if current_check_itn.strip() == "" or current_check_itn in list_of_unique_itn:
                    continue
                else:
                    # блок добавляет немного магии в сбор данных
                    # кол-во машин у одного контрагента (при условии сортировки по контрагенту)
                    # найдем из длины списка машин

                    # соберем список машин для одного контрагента
                    vehicle_number_list = []
                    # делаем собиратель статусов ( если хоть один "да" - общий статус "да") для аренды и т.п.
                    sign_list = [False, False, False, False]

                    temp_itn = current_check_itn
                    temp_index_row = i_row

                    while temp_itn == current_check_itn and temp_itn.strip() != "":
                        # собираем список машин каждого контрагента

                        vehicle_number_list.append(str(active_sheet.cell(row=temp_index_row,
                                                                         column=oc_num["Номер ТС"][0]).value).strip())
                        # хитрый механизм сбора признаков - двумерная структура сначала проходит по строке,
                        # а затем добавляет элемент след строки, пока не станет True
                        for sign_id in range(len(sign_list_names)):
                            if sign_list[sign_id] is False:
                                if str(active_sheet.cell(row=temp_index_row,
                                                         column=oc_num[sign_list_names[sign_id]][0]).value)[0] == "Д":
                                    sign_list[sign_id] = True

                        temp_index_row += 1
                        temp_itn = str(
                            active_sheet.cell(row=temp_index_row, column=oc_num["ИНН"]).value).strip()
                    # теперь у нас есть список машин и список признаков для одного и того же ИНН

                    # пожалуй, тут же будем собирать остальные данные, требующие обработки,
                    # чтобы потом блоком залить их FIXME

                    # определяем наиболее часто втречающийся регион по номеру машины
                    region_vehicle = other_function_xlsx.region_on_number_vehicle(vehicle_number_list)
                    region_vehicle_name = ""
                    # готовим регион ТС в строку для записи
                    if region_vehicle in fab.RegionDistrict.vehicle_code_dict:
                        region_vehicle_name = fab.RegionDistrict.vehicle_code_dict[region_vehicle]

                    # определяем регион по адресам (приоритет фактическому адресу)
                    fact_address = str(active_sheet.cell(row=i_row, column=oc_num["Факт. адрес"]).value)
                    f_country, f_region_name, f_district_name, f_city_name, f_sign_good_address = \
                        fab.RegionDistrict.find_address(fact_address)
                    legal_address = str(active_sheet.cell(row=i_row, column=oc_num["Юр. адрес"]).value)
                    l_country, l_region_name, l_district_name, l_city_name, l_sign_good_address = \
                        fab.RegionDistrict.find_address(legal_address)
                    if f_sign_good_address is True:
                        good_country = f_country
                        good_region_name = f_region_name
                        good_district_name = f_district_name
                        sign_good_address = True
                    elif l_sign_good_address is True:
                        good_country = l_country
                        good_region_name = l_region_name
                        good_district_name = l_district_name
                        sign_good_address = True


                    if ("Беларусь" in fact_address or "Беларусь" in legal_address) and sign_good_address is False:
                        good_country = "Беларусь"
                        good_region_name = ""
                        good_district_name = ""
                        sign_good_address = f_sign_good_address + l_sign_good_address
                    elif ("Казахстан" in fact_address or "Казахстан" in legal_address) and sign_good_address is False:
                        good_country = "Казахстан"
                        good_region_name = ""
                        good_district_name = ""
                        sign_good_address = f_sign_good_address + l_sign_good_address
                    else:
                        good_country = f_country if (f_country is not None and f_country != "") else l_country
                        good_region_name = f_region_name if f_region_name is not None else l_region_name
                        good_district_name = f_district_name if f_district_name is not None else l_district_name
                        sign_good_address = f_sign_good_address + l_sign_good_address

                    # блок работы с телефонами
                    # находим сотовые телефоны и приводим их к стандартному виду
                    phone_numbers_string = str(active_sheet.cell(row=i_row, column=oc_num["Телефон"]).value)
                    # получаем список мобильных телефонов компании
                    company_phone_list = phone_func.make_good_phone_list(phone_numbers_string)

                    # делаем ячейки с телефонами
                    sms_phone = ""
                    extended_phones = ""
                    if len(company_phone_list) > 0:
                        sms_phone = company_phone_list[0]
                        if len(company_phone_list) > 1:
                            extended_phones = ", ".join(company_phone_list[1:])

                    list_of_unique_itn.append(current_check_itn)

                    # продолжаем собирать элементы новой строки
                    row_for_write = [""] * 26
                    # блок распределения уже собранных данных
                    for num in range(len(sign_list_names)):
                        row_for_write[oc_num[sign_list_names[num]][1]] = "Да" if sign_list[num] else "Нет"
                    # print(vehicle_number_list)
                    row_for_write[oc_num["Номер ТС"][1]] = ", ".join(vehicle_number_list)
                    row_for_write[new_num["Кол-во ТС"]] = len(vehicle_number_list)
                    # прописываем регион ТС в строку для записи
                    row_for_write[new_num["Регион ТС"]] = region_vehicle_name
                    row_for_write[new_num["СМС"]] = sms_phone
                    row_for_write[new_num["Доп. номера"]] = extended_phones
                    row_for_write[new_num["Страна"]] = good_country
                    row_for_write[new_num["Округ"]] = good_district_name
                    if good_region_name is not None:
                        row_for_write[new_num["Регион"]] = good_region_name.split(" / ")[0]
                    else:
                        row_for_write[new_num["Регион"]] = good_region_name
                    row_for_write[new_num["Юр. город"]] = l_city_name
                    row_for_write[new_num["Факт. город"]] = f_city_name
                    row_for_write[new_num["Адрес не распознан"]] = "" if sign_good_address else "Да"

                    # заполняем оставшиеся ячейки, из анных которым нужен просто перенос
                    for j_col in range(1, numbers_cols + 1):
                        # собираем новую строку, исключая ненужные столбцы
                        if j_col not in ignore_cols_list_original_head_id:
                            current_cell_value = str(active_sheet.cell(row=i_row, column=j_col).value)
                            # ячейки, данные которых нужно просто перенести - переносим в соответствующие ячейки
                            # согласно словарю простого переноса
                            if j_col in just_transfer_list:
                                row_for_write[just_transfer_list[j_col]] = \
                                    current_cell_value if (
                                                                  current_cell_value is not None) or current_cell_value != "None" else ""

                    # чистим от None перед записью
                    for cell_id in range(len(row_for_write)):
                        if row_for_write[cell_id] is None or row_for_write[cell_id] == "None":
                            row_for_write[cell_id] = ""

                    # определим лист, куда будем записывать данные, в зависимости от округа good_district_name
                    counter_processed_rows += 1
                    counter_from_file += 1
                    true_address = "Беларусь" not in fact_address and "Беларусь" not in legal_address and \
                                   "Казахстан" not in fact_address and "Казахстан" not in legal_address and \
                                   good_country == fab.RegionDistrict.country


                    if true_address:
                        for district_id in range(len(fab.RegionDistrict.district_tuple)):
                            if fab.RegionDistrict.district_tuple[district_id] == good_district_name:
                                exit_phone_xlsx.active = district_id
                                exit_phone_xlsx.active.append(row_for_write)
                                break
                    # если не нашли округ, пишем в ненайденные
                    else:
                        exit_phone_xlsx.active = len(fab.RegionDistrict.district_tuple) - 1
                        exit_phone_xlsx.active.append(row_for_write)
            current_xlsx_obj.close()
            print(f"Файл обработан и закрыт. Уникальных контрагентов в файле: "
                  f"{str(counter_from_file)} из {str(numbers_rows)} записей")
            print("_" * 20)

    exit_phone_xlsx.active = 0
    print('Производится запись на диск')
    exit_phone_xlsx.save(co_da.dir_output_files + "\\" + co_da.exit_phone_base_file_name)
    print(f"В файле {co_da.exit_phone_base_file_name} сохранены изменения")

    exit_phone_xlsx.close()
    print("Файл закрыт")
    print("Обработано уникальных контрагентов", counter_processed_rows)
    print("Программа успешно завершена.")

    # input('Для заверения нажмите Enter...')

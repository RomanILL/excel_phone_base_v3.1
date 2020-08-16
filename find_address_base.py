import openpyxl
import func_xlsx_base as fxb
import config_data


class RegionDistrict:
    """
    класс определяющий принадлежность к федеральному округу по строке
    синглтон по сути, но пока реализован как обычный
    """
    # блок входных данных
    country = "Russia"
    support_dir = "support_tables"
    address_file = "address_structure.xlsx"
    address_base_file = support_dir + "\\" + address_file
    district_tuple = ("Центральный", "Приволжский", "Дальневосточный", "Северо-Западный",
                      "Северо-Кавказский", "Сибирский", "Уральский", "Южный")
    region_name_tuple = None  # чуть позже создадим кортеж из списка для ускорения процессов
    post_index_dict = dict()
    region_base = dict()
    region_name_list = list()
    cities_name_dict = dict()

    def __init__(self):

        self.country_name = RegionDistrict.country
        RegionDistrict.make_region_base()

    @classmethod
    def make_region_base(cls):
        # в каких столбиках какие данные
        head_district_number = 5
        head_region_name = 2
        head_post_index = 3
        head_vehicle_code = 4

        # формируем структуру данных для адресов
        # открываем файл со структурой
        with openpyxl.load_workbook(RegionDistrict.address_base_file) as district_file:
            district_file.active = 0
            # перебираем файл, составляем рабочие списки и словари
            for i in range(2, district_file.active.max_row + 1):
                this_post_index = district_file.active.cell(row=i, column=head_post_index)
                this_vehicle_code = district_file.active.cell(row=i, column=head_vehicle_code)
                this_district_number = district_file.active.cell(row=i, column=head_district_number)
                this_region_name = district_file.active.cell(row=i, column=head_region_name)
                # заполняем словарь класса с регионами
                cls.region_base[i - 2] = this_region_name, this_district_number, this_vehicle_code, this_post_index
                # FIXME из предыдущей строки нужно убрать что не будем использовать
                #  первые кандидаты this_vehicle_code и this_post_index, так как по этим полям идет поиск региона
                # заодно заполняем словарь почтовых индексов
                cls.post_index_dict[this_post_index] = i - 2
            # переключаемся на вкладку с названиями регионов и собираем список названий
            # (используем для принадлежности городов)
            district_file.active += 1
            for i in range(2, district_file.active.max_row + 1):
                cls.region_name_list.append(district_file.active.cell(row=i, column=1))
                # Заодно создадим словарь регионов с пустыми списками городов
                cls.cities_name_dict[i - 2] = set()
            # переключаемся на вкладку с названиями городов и собираем словарь регионов со списками городов
            district_file.active += 1

            for i in range(2, district_file.active.max_row + 1):
                this_city = district_file.active.cell(row=i, column=1)
                this_region_id = district_file.active.cell(row=i, column=2)
                # this_district_id = district_file.active.cell(row=i, column=3)
                cls.cities_name_dict[this_region_id].add(this_city)
        cls.region_name_tuple = tuple(cls.region_name_list)

    @classmethod
    def get_region_and_district_on_region_id(cls, region_id):
        region_name = cls.region_base[region_id][0]
        district_number = cls.region_base[region_id][1]
        district_name = cls.district_tuple[district_number]

        return region_name, district_name

    @classmethod
    def find_address(cls, any_address_string):
        flag = 0  # пока регион не найден
        # сначала будем искать по почтовому индексу - наййдем индекс в строке
        post_index = cls.search_post_index(any_address_string)
        region_id = cls.post_index_dict.get(post_index)
        if not region_id is None:
            # если почтовый индекс нашли, запоминаем имя региона и федерального округа
            region_name, district_name = cls.get_region_and_district_on_region_id(region_id)

            flag = 1  # нехватает только названия города
        # если по почтовому индексу ничего нет (или не указан индекс) будем искать по имени области
        elif flag == 0:

            region_id = cls.searsh_region_id(any_address_string)
            region_name, district_name = cls.get_region_and_district_on_region_id(region_id)

            flag = 1  # нехватает только названия города
        elif flag == 0:
            # если и по области ничего нет - надо искать по названию города (как повезет)
            region_id, city_name = cls.searh_region_on_city_name(any_address_string)
            region_name, district_name = cls.get_region_and_district_on_region_id(region_id)

            flag = 2  # нашли про адрес округ, регион и город
        elif flag == 1:
            # будем искать город в регионе
            city_name = cls.searsh_city_name_in_region(any_address_string, region_id)
            flag = 2
        if flag == 2:
            return cls.country, region_name, district_name, city_name
        elif flag == 1:
            return cls.country, region_name, district_name, None
        elif flag == 0:
            return None, None, None, None

    def search_post_index(self, any_address_string):
        for i in range(len(any_address_string) - 6):
            if any_address_string[i: i + 6].isdigit():
                return any_address_string[i: i + 3]

    @classmethod
    def searsh_region_id(cls, any_address_string):
        for current_id in range(len(cls.region_name_tuple)):
            for region_name_up in cls.region_name_tuple[current_id].split("/").strip().upper():
                if region_name_up in any_address_string.upper():
                    region_id = current_id
                    return region_id

    @classmethod
    def searh_region_on_city_name(cls, any_address_string):
        for current_region_id, city_names_of_region in cls.cities_name_dict.items():
            for current_city_name in city_names_of_region:
                if current_city_name.upper() in any_address_string.upper():
                    region_id = current_region_id
                    city_name = current_city_name
                    return region_id, city_name

    @classmethod
    def searh_city_name_in_region(cls, any_address_string, current_region_id):
        city_names_of_region = cls.cities_name_dict.items(current_region_id)
        for current_city_name in city_names_of_region:
            if current_city_name.upper() in any_address_string.upper():
                return current_city_name

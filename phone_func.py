import re

def make_good_phone_list(phone_string):
    # делим строку с телефонами на отдельные телефоны
    delimiters = ".", ",", ":", ";", "\\", "/", "+", "или", "и"
    regexPattern = "|".join(map(re.escape, delimiters))
    mobile_phone_candidate_list = re.split(regexPattern, phone_string)

    # чистим от мусора и городских телефонов
    rus_standard_mobile_list = _make_mobile_list(mobile_phone_candidate_list)

    return rus_standard_mobile_list


def _make_mobile_list(origin_phone_list):
    good_phone_list = list()
    # перебираем номера из оригинального списка
    for phone_candidate in origin_phone_list:
        full_number = ""

        for number_x in phone_candidate:
            if number_x in "0123456789":
                full_number += number_x

        if len(full_number) == 11 and (full_number[:2] == "79" or full_number[:2] == "89"):
            full_number = _make_standard_rus(full_number)
            good_phone_list.append(full_number)

        elif len(full_number) == 10 and full_number[:1] == "9":
            full_number = "7" + full_number
            full_number = _make_standard_rus(full_number)
            good_phone_list.append(full_number)

    return good_phone_list


def _make_standard_rus(number_11):
    """ 79xx1234567 -> +7 (9xx) 1234567 """
    number_11 = f"+7 ({number_11[1:4]}) {number_11[4:]}"
    return number_11
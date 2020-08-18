import os


# глобальные переменные и настройки программы для разных модулей

def check_destination_folders(folders_list):
    for folder in folders_list:
        check_and_create_dir(folder)

    print("-" * 30 + "\nПроверка успешно выполнена")


def check_and_create_dir(dir_name):
    """функция проверяет существует ли папка, и если её нет, то создает новую"""
    if not os.path.isdir(dir_name):
        os.mkdir(dir_name)
        return f"Папка '{dir_name}' не существовала. Папка создана"
    return f"Папка '{dir_name}' существует"


if __name__ == "__main__":
    print("Этот модуль так не используется - только данные настроек для импорта")

if __name__ != "__main__":
    version = 3.1
    exit_phone_base_file_name = f"global_base_v{version}.xlsx"
    dir_input_files = "input_xlsx"
    dir_output_files = "output_xlsx"
    dir_supporting_files = "support_tables"
    print("Версия программы", version)
    print("Проверка целостности системы каталогов")
    check_destination_folders((dir_input_files, dir_output_files, dir_supporting_files))

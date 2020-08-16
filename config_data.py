# глобальные переменные и настройки программы для разных модулей
import inspect
if __name__ == "__main__":
    print(__module__)
    y  = "dkd"
if __name__ != "__main__":
    x = "Вот это поворот!"
    print(__name__)
    print(x)
    print(inspect.stack())
"""Валидаторы для внутренних функций."""
import datetime
import os
import sys

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
elif __file__:
    BASE_DIR = os.path.dirname(__file__)


class ValidateError(Exception):
    pass


def now():
    return datetime.datetime.now().strftime("%d.%m %H_%M_%S")


def log_any_error(any_text):
    """Функция для записи ошибок в файл."""
    log = open(rf"{BASE_DIR}\log_any_error.txt", "a+")
    text = str(any_text)
    print(f"[{now()}] {text}", file=log)
    log.close()


def validate_column_with_data(row: list, expected_values: int) -> bool:
    """Пропускает строку, если количество "None" в списке/строке больше ожидаемого."""
    count_values = len(row) - row.count(None)
    if count_values < expected_values:
        return True
    return False


def validate_department(department) -> bool:
    """Проверка поля отделения: более 1 символа и не None."""
    if department is None:
        return False
    if len(department) > 1:
        return True
    else:
        return False


def validate_count_days(count_days) -> int:
    """Возвращает количество дней или ноль."""
    if count_days is None:
        return 0
    try:
        return int(count_days)
    except ValueError:
        log_any_error("Не корректное значение количества дней в отчете!")
        return 0


def validate_number_history(number_history):
    """Проверяем наличие номера истории и возвращаем его."""
    if number_history is None:
        log_any_error("Не корректное значение номера карты!")
    return number_history


def validate_bunks_from_file(value, department, BUNKS):
    """Проверяем правильно ли введены данные в файл коечного фонда."""
    try:
        count_bunks = value if type(value) == int else int(value.strip())
    except ValueError:
        log_any_error(
            f"[ERR] Для отделения <{department}> указано не числовое"
            f" значение количества коек <{value}>"
        )
    except TypeError:
        log_any_error(
            "Значение количества дней не число. Проверьте, что ввели число!"
            f"Число дней в файле с отделениями и койками = {value}"
        )
    else:
        BUNKS[department] = count_bunks


def validate_for_title(target_values: list, row: list):
    """Является ли строка заголовком."""
    if len(target_values) == 1:
        search = target_values[0].upper()
        if search in set([element.upper() if type(element) is str else element for element in row if element is not None]):
            return True
        return False
    
    values = [value.upper() if type(value) is str else value for value in target_values]
    for value in values:
        if value in set([element.upper() if type(element) is str else element for element in row if element is not None]):
            return True
        return False


def validate_numbers(row: list):
    """Если часть значений строка с номерами столбцев."""
    lst = set(['1', '2',])
    if set([element for element in row[:2] if element is not None]) == lst:
        return True
    return False


def validate_not_pdo(target_values: list, row: str):
    """Если есть ПДО, то возвращаем ложь, иначе истину."""
    values = [value.upper() if type(value) is str else value for value in target_values]
    for value in values:
        if value in row.upper():
            return False
        return True
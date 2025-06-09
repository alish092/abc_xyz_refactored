from enum import Enum

class StandardColumns(str, Enum):
    ARTIKUL = "Артикул"
    NOMENCLATURA = "Номенклатура"
    SUMMA = "Сумма"           # Стандартное название для денежных значений
    OSTATOK = "Остаток"
    ABC = "ABC"
    XYZ = "XYZ"
    DOLIA = "Доля"
    NAKOPITEL = "Накопл"
    CV = "CV"
    MESYAC = "Месяц"

class SourceColumns(str, Enum):
    # Исходные названия из файлов
    VIRUCHKA = "Выручка"      # Из файла продаж
    KOLICHESTVO = "Количество" # Из файла остатков
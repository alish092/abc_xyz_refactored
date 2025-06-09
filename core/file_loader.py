import pandas as pd
from core.data_normalizer import DataNormalizer
from config.column_schema import StandardColumns


def load_stock(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, header=4)

    # Проверяем исходные колонки
    required = ['Артикул', 'Номенклатура', 'Количество']
    for col in required:
        if col not in df.columns:
            raise ValueError(f"[STOCK] Ожидаются колонки: {required}")

    # Обрабатываем данные
    df['Артикул'] = df['Артикул'].astype(str).str.strip().str.lower()
    df['Количество'] = pd.to_numeric(df['Количество'], errors='coerce').fillna(0)

    # Нормализуем к стандартным названиям
    df = DataNormalizer.normalize_stock(df)

    return df[[StandardColumns.ARTIKUL, StandardColumns.NOMENCLATURA, StandardColumns.OSTATOK]]
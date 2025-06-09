import pandas as pd
from config.column_schema import StandardColumns, SourceColumns


class DataNormalizer:
    """Приводит данные к стандартным названиям колонок"""

    @staticmethod
    def normalize_sales(df: pd.DataFrame) -> pd.DataFrame:
        """Нормализует данные продаж к стандартному виду"""
        df = df.copy()

        # Переименовываем колонки к стандартным названиям
        column_mapping = {
            SourceColumns.VIRUCHKA: StandardColumns.SUMMA
        }

        df = df.rename(columns=column_mapping)
        return df

    @staticmethod
    def normalize_stock(df: pd.DataFrame) -> pd.DataFrame:
        """Нормализует данные остатков к стандартному виду"""
        df = df.copy()

        # Переименовываем колонки к стандартным названиям
        column_mapping = {
            SourceColumns.KOLICHESTVO: StandardColumns.OSTATOK
        }

        df = df.rename(columns=column_mapping)
        return df

    @staticmethod
    def validate_required_columns(df: pd.DataFrame, required_columns: list) -> None:
        """Проверяет наличие обязательных колонок"""
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            raise ValueError(f"Отсутствуют обязательные колонки: {missing}")
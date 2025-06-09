import numpy as np
import pandas as pd
from config.schema import Thresholds
from config.column_schema import StandardColumns
from core.data_normalizer import DataNormalizer


class ABCAnalyzer:
    def __init__(self, thresholds: Thresholds):
        self.thresholds = thresholds

    def analyze(self, df: pd.DataFrame) -> pd.DataFrame:
        # Валидация входных данных
        DataNormalizer.validate_required_columns(
            df, [StandardColumns.SUMMA]
        )

        df = df.copy()
        total = df[StandardColumns.SUMMA].sum()
        df[StandardColumns.DOLIA] = df[StandardColumns.SUMMA] / total
        df[StandardColumns.NAKOPITEL] = df[StandardColumns.DOLIA].cumsum()

        condlist = [
            df[StandardColumns.NAKOPITEL] <= self.thresholds.A,
            df[StandardColumns.NAKOPITEL] <= self.thresholds.B
        ]
        choicelist = ['A', 'B']
        df[StandardColumns.ABC] = np.select(condlist, choicelist, default='C')
        return df


class XYZAnalyzer:
    def __init__(self, thresholds: Thresholds):
        self.thresholds = thresholds

    def analyze(self, df: pd.DataFrame) -> pd.DataFrame:
        # Валидация входных данных
        required_stats = ['mean', 'std', 'count']
        DataNormalizer.validate_required_columns(df, required_stats)

        df = df.copy()
        df[StandardColumns.CV] = np.where(df['count'] > 1, df['std'] / df['mean'], 0)

        def classify(cv: float) -> str:
            if pd.isna(cv) or cv == np.inf:
                return 'Z'
            if cv <= self.thresholds.X:
                return 'X'
            elif cv <= self.thresholds.Y:
                return 'Y'
            return 'Z'

        df[StandardColumns.XYZ] = df[StandardColumns.CV].apply(classify)
        return df
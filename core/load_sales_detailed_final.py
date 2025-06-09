
import pandas as pd
from config.column_schema import StandardColumns
from core.data_normalizer import DataNormalizer

def load_sales_detailed(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, header=None)

    records = []
    current_sklad = None
    current_month = None

    for i in range(1, len(df)):
        row = df.iloc[i]
        cell = str(row[0]).strip()

        if 'склад' in cell.lower():
            current_sklad = cell
            continue

        if any(m in cell.lower() for m in ('янв', 'фев', 'мар', 'апр', 'май', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек')):
            current_month = cell
            continue

        if ',' in cell:
            parts = cell.split(',', 1)
            artikul = parts[0].strip().lower()
            name = parts[1].strip()

            try:
                qty = float(str(row[1]).replace(" ", "").replace(",", "."))
            except:
                qty = 0

            try:
                revenue = float(str(row[2]).replace(" ", "").replace(",", "."))
            except:
                revenue = 0

            records.append({
                StandardColumns.SKLAD: current_sklad,
                StandardColumns.MONTH: current_month,
                StandardColumns.ARTICLE: artikul,
                StandardColumns.NAME: name,
                StandardColumns.QTY: qty,
                StandardColumns.REVENUE: revenue,
            })

    df = pd.DataFrame(records)
    df = DataNormalizer.normalize_sales(df)
    return df

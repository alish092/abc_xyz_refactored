
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
from config.column_schema import StandardColumns
from core.file_loader import load_stock
from core.load_sales_detailed import load_sales_detailed
from core.analyzer import ABCAnalyzer, XYZAnalyzer
from config.schema import AppConfig

class AppGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title('ABC/XYZ-анализатор')
        self.root.geometry('1000x600')

        self.sales_path = None
        self.stock_path = None
        self.df = pd.DataFrame()

        self.config = AppConfig()
        self.abc_analyzer = ABCAnalyzer(self.config.thresholds)
        self.xyz_analyzer = XYZAnalyzer(self.config.thresholds)

        self.setup_widgets()

    def sort_by_column(self, col):
        reverse = getattr(self, 'sort_reverse', {}).get(col, False)
        self.df = self.df.sort_values(by=col, ascending=not reverse)
        self.sort_reverse = {col: not reverse}
        self.update_table()

    def setup_widgets(self):
        top_frame = tk.Frame(self.root)
        top_frame.pack(pady=10)
        tk.Button(top_frame, text='💾 Сохранить в Excel', command=self.save_to_excel).grid(row=0, column=3, padx=10)

        tk.Button(top_frame, text='Загрузить продажи', command=self.load_sales_file).grid(row=0, column=0, padx=10)
        tk.Button(top_frame, text='Загрузить остатки', command=self.load_stock_file).grid(row=0, column=1, padx=10)
        tk.Button(top_frame, text='Рассчитать', command=self.try_analyze).grid(row=0, column=2, padx=10)

        self.status_label = tk.Label(self.root, text='📁 Загрузите файлы для анализа')
        self.status_label.pack()

        tree_frame = tk.Frame(self.root)
        tree_frame.pack(expand=True, fill='both')

        scrollbar = ttk.Scrollbar(tree_frame, orient='vertical')
        scrollbar.pack(side='right', fill='y')

        self.tree = ttk.Treeview(
            tree_frame,
            columns=['Артикул', 'Номенклатура', 'Сумма', 'ABC', 'XYZ', 'Остаток', 'ABC_XYZ', 'Рекомендация'],
            show='headings',
            yscrollcommand=scrollbar.set
        )
        scrollbar.config(command=self.tree.yview)

        for col in ['Артикул', 'Номенклатура', 'Сумма', 'ABC', 'XYZ', 'Остаток', 'ABC_XYZ', 'Рекомендация']:
            self.tree.heading(col, text=col, command=lambda _col=col: self.sort_by_column(_col))
            self.tree.column(col, width=130)

        self.tree.pack(expand=True, fill='both')
        self.tree.bind("<Control-c>", self.copy_selection_to_clipboard)
        self.tree.bind("<ButtonRelease-1>", lambda e: self.tree.focus_set())  # для фокуса

    def copy_selection_to_clipboard(self, event=None):
        selection = self.tree.selection()
        if not selection:
            return

        rows = []
        for item in selection:
            values = self.tree.item(item)['values']
            rows.append('\t'.join(str(v) for v in values))

        self.root.clipboard_clear()
        self.root.clipboard_append('\n'.join(rows))
        self.root.update()  # необходимо для некоторых ОС

    def load_sales_file(self):
        path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')])
        if path:
            try:
                load_sales_detailed(path)
                self.sales_path = path
                self.status_label.config(text='✅ Продажи загружены')
            except Exception as e:
                messagebox.showerror('Ошибка загрузки продаж', str(e))

    def load_stock_file(self):
        path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')])
        if path:
            try:
                load_stock(path)
                self.stock_path = path
                self.status_label.config(text='✅ Остатки загружены')
            except Exception as e:
                messagebox.showerror('Ошибка загрузки остатков', str(e))

    def try_analyze(self):
        if not self.sales_path or not self.stock_path:
            return

        try:
            sales = load_sales_detailed(self.sales_path)
            stock = load_stock(self.stock_path)

            abc_df = sales.groupby([
                StandardColumns.ARTIKUL,
                StandardColumns.NOMENCLATURA
            ])[StandardColumns.SUMMA].sum().reset_index()
            abc_df = self.abc_analyzer.analyze(abc_df)

            stats = sales.groupby([
                StandardColumns.ARTIKUL,
                StandardColumns.NOMENCLATURA
            ])[StandardColumns.SUMMA].agg(['mean', 'std', 'count']).reset_index()
            xyz_df = self.xyz_analyzer.analyze(stats)

            df = pd.merge(
                abc_df,
                xyz_df[[StandardColumns.ARTIKUL, StandardColumns.NOMENCLATURA, StandardColumns.XYZ]],
                on=[StandardColumns.ARTIKUL, StandardColumns.NOMENCLATURA],
                how='left'
            )
            df = pd.merge(
                df,
                stock[[StandardColumns.ARTIKUL, StandardColumns.OSTATOK]],
                on=StandardColumns.ARTIKUL,
                how='left'
                )
            df[StandardColumns.OSTATOK] = df[StandardColumns.OSTATOK].fillna(0)

            self.df = df
                    self.df['ABC_XYZ'] = self.df[StandardColumns.ABC] + self.df[StandardColumns.XYZ]
        def get_recommendation(code):
            if code == 'AX': return 'Поддерживать высокий запас, ключевой товар'
            elif code in ['AY', 'AZ']: return 'Контролировать запасы, прогнозировать спрос'
            elif code == 'BX': return 'Поддерживать, но оптимизировать'
            elif code in ['BY', 'BZ']: return 'Периодический пересмотр, анализ сезонности'
            elif code == 'CX': return 'Снизить остатки, невостребованный товар'
            elif code in ['CY', 'CZ']: return 'Рассмотреть уценку или списание'
            else: return 'Н/Д'
        self.df['Рекомендация'] = self.df['ABC_XYZ'].apply(get_recommendation)
self.update_table()
            self.status_label.config(text=f'✅ Готово: {len(df)} позиций')

        except Exception as e:
            messagebox.showerror('Ошибка анализа', str(e))

    def update_table(self):
        self.tree.delete(*self.tree.get_children())
        for _, row in self.df.iterrows():
            values = [
                row[StandardColumns.ARTIKUL],
                row[StandardColumns.NOMENCLATURA],
                f"{row[StandardColumns.SUMMA]:.0f}",
                row[StandardColumns.ABC],
                row[StandardColumns.XYZ],
                f"{row[StandardColumns.OSTATOK]:.0f}"
            ]
            self.tree.insert('', 'end', values=values)

    def save_to_excel(self):
        if self.df.empty:
            messagebox.showinfo("Нет данных", "Сначала выполните анализ.")
            return

        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel файлы", "*.xlsx")])
        if path:
            try:
                self.df.to_excel(path, index=False)
                messagebox.showinfo("Успешно", f"Данные сохранены в файл:\n{path}")
            except Exception as e:
                messagebox.showerror("Ошибка сохранения", str(e))

    def run(self):
        self.root.mainloop()

def run_app():
    AppGUI().run()

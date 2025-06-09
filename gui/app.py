
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

        # Первая строка кнопок
        tk.Button(top_frame, text='Загрузить продажи', command=self.load_sales_file).grid(row=0, column=0, padx=10)
        tk.Button(top_frame, text='Загрузить остатки', command=self.load_stock_file).grid(row=0, column=1, padx=10)
        tk.Button(top_frame, text='Рассчитать', command=self.try_analyze).grid(row=0, column=2, padx=10)
        tk.Button(top_frame, text='💾 Сохранить в Excel', command=self.save_to_excel).grid(row=0, column=3, padx=10)

        # Строка поиска
        search_frame = tk.Frame(self.root)
        search_frame.pack(pady=(0, 10))

        tk.Label(search_frame, text='🔍 Поиск:').pack(side='left', padx=(0, 5))
        self.search_entry = tk.Entry(search_frame, width=30)
        self.search_entry.pack(side='left', padx=(0, 10))
        self.search_entry.bind('<KeyRelease>', self.on_search)

        tk.Button(search_frame, text='Очистить', command=self.clear_search).pack(side='left', padx=(0, 10))

        # Счетчик найденных записей
        self.search_label = tk.Label(search_frame, text='')
        self.search_label.pack(side='left')

        self.status_label = tk.Label(self.root, text='📁 Загрузите файлы для анализа')
        self.status_label.pack()

        tree_frame = tk.Frame(self.root)
        tree_frame.pack(expand=True, fill='both')

        scrollbar = ttk.Scrollbar(tree_frame, orient='vertical')
        scrollbar.pack(side='right', fill='y')
        h_scrollbar = ttk.Scrollbar(tree_frame, orient='horizontal')
        h_scrollbar.pack(side='bottom', fill='x')

        self.tree = ttk.Treeview(
            tree_frame,
            columns=['Артикул', 'Номенклатура', 'Сумма', 'ABC', 'XYZ', 'Остаток', 'ABC_XYZ', 'Рекомендация'],
            show='headings',
            yscrollcommand=scrollbar.set,
            xscrollcommand=h_scrollbar.set,
            selectmode='none'
        )
        scrollbar.config(command=self.tree.yview)
        h_scrollbar.config(command=self.tree.xview)

        for col in ['Артикул', 'Номенклатура', 'Сумма', 'ABC', 'XYZ', 'Остаток', 'ABC_XYZ', 'Рекомендация']:
            self.tree.heading(col, text=col, command=lambda _col=col: self.sort_by_column(_col))
            self.tree.column(col, width=130)

        self.tree.pack(expand=True, fill='both')

        # Настраиваем стили для выделения
        style = ttk.Style()
        style.map("Treeview",
                  background=[('selected', ''), ('active', '')],
                  foreground=[('selected', ''), ('active', '')])

        # Привязываем события
        self.tree.bind("<Button-1>", self.on_cell_click)
        self.tree.bind("<Control-c>", self.copy_cell_to_clipboard)
        self.tree.bind("<Escape>", self.clear_cell_selection)  # Добавляем Escape для очистки
        self.tree.bind("<Button-3>", self.clear_cell_selection)  # Правая кнопка мыши для очистки

        # Переменные для отслеживания выделенной ячейки
        self.selected_item = None
        self.selected_column = None
        self.selection_tags = []  # Для отслеживания тегов выделения
        self.highlighted_items = []  # Для отслеживания подсвеченных элементов

        # Сохраняем оригинальные данные для поиска
        self.original_df = pd.DataFrame()
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="📋 Копировать ячейку", command=self.copy_cell_to_clipboard)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="❌ Снять выделение", command=self.clear_cell_selection)

        # Привязываем правую кнопку мыши к контекстному меню
        self.tree.bind("<Button-3>", self.show_context_menu)

    def show_context_menu(self, event):
        """Показывает контекстное меню"""
        # Сначала выделяем ячейку под курсором
        self.on_cell_click(event)

        # Показываем меню
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()
            
    def on_cell_click(self, event):
        """Определяет и выделяет ячейку, на которую кликнул пользователь"""
        # Очищаем предыдущее выделение
        self.clear_cell_selection()

        item = self.tree.identify('item', event.x, event.y)
        column = self.tree.identify('column', event.x, event.y)

        if item and column:
            self.selected_item = item
            self.selected_column = column

            # Создаем визуальное выделение ячейки
            self.highlight_cell(item, column)

            # Получаем информацию о ячейке
            col_names = ['Артикул', 'Номенклатура', 'Сумма', 'ABC', 'XYZ', 'Остаток', 'ABC_XYZ', 'Рекомендация']
            col_index = int(column.replace('#', '')) - 1

            if 0 <= col_index < len(col_names):
                col_name = col_names[col_index]
                cell_value = self.tree.item(item)['values'][col_index]

                # Обновляем статус
                current_status = self.status_label.cget('text')
                if '|' in current_status:
                    base_status = current_status.split('|')[0].strip()
                else:
                    base_status = current_status

                self.status_label.config(text=f'{base_status} | 📍 Выделено: {col_name} = {cell_value} (Escape - снять выделение)')

    def highlight_cell(self, item, column):
        """Визуально выделяет конкретную ячейку используя теги"""
        # Создаем уникальный тег для этой ячейки
        tag_name = f"selected_{item}_{column}"

        # Настраиваем стиль для выделенной ячейки
        self.tree.tag_configure(tag_name, background='lightblue', foreground='black')

        # Применяем тег к элементу
        current_tags = list(self.tree.item(item, 'tags'))
        current_tags.append(tag_name)
        self.tree.item(item, tags=current_tags)

        # Сохраняем информацию о теге для последующего удаления
        self.selection_tags.append((item, tag_name))

    def clear_cell_selection(self, event=None):
        """Очищает визуальное выделение ячеек"""
        # Удаляем все теги выделения
        for item, tag_name in self.selection_tags:
            try:
                current_tags = list(self.tree.item(item, 'tags'))
                if tag_name in current_tags:
                    current_tags.remove(tag_name)
                    self.tree.item(item, tags=current_tags)
            except tk.TclError:
                # Элемент мог быть удален при обновлении таблицы
                pass

        self.selection_tags.clear()
        self.selected_item = None
        self.selected_column = None

        # Обновляем статус, убирая информацию о выделении
        current_status = self.status_label.cget('text')
        if '|' in current_status:
            base_status = current_status.split('|')[0].strip()
            self.status_label.config(text=base_status)

    def copy_cell_to_clipboard(self, event=None):
        """Копирует содержимое выделенной ячейки в буфер обмена"""
        if self.selected_item and self.selected_column:
            try:
                col_index = int(self.selected_column.replace('#', '')) - 1
                values = self.tree.item(self.selected_item)['values']

                if 0 <= col_index < len(values):
                    cell_value = str(values[col_index])

                    # Очищаем буфер обмена и добавляем значение
                    self.root.clipboard_clear()
                    self.root.clipboard_append(cell_value)

                    # Принудительно обновляем буфер обмена
                    try:
                        self.root.update_idletasks()
                        self.root.update()
                    except:
                        pass

                    # Показываем уведомление
                    col_names = ['Артикул', 'Номенклатура', 'Сумма', 'ABC', 'XYZ', 'Остаток', 'ABC_XYZ', 'Рекомендация']
                    col_name = col_names[col_index] if col_index < len(col_names) else 'Значение'

                    # Временно меняем статус для показа уведомления
                    original_status = self.status_label.cget('text')
                    self.status_label.config(text=f'📋 Скопировано: {cell_value}')
                    self.root.after(2000, lambda: self.status_label.config(text=original_status))

            except Exception as e:
                # Если не удалось скопировать, показываем ошибку
                original_status = self.status_label.cget('text')
                self.status_label.config(text=f'❌ Ошибка копирования: {str(e)}')
                self.root.after(3000, lambda: self.status_label.config(text=original_status))
        else:
            # Если ничего не выделено, показываем подсказку
            original_status = self.status_label.cget('text')
            self.status_label.config(text='❌ Сначала кликните на ячейку для выделения')
            self.root.after(2000, lambda: self.status_label.config(text=original_status))

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

    def on_search(self, event=None):
        """Фильтрует данные по поисковому запросу"""
        if self.original_df.empty:
            return

        query = self.search_entry.get().lower().strip()

        if not query:
            # Если поисковая строка пуста - показываем все данные
            self.df = self.original_df.copy()
            self.update_table()
            self.search_label.config(text='')
            return

        # Фильтруем данные
        mask = (
                self.original_df[StandardColumns.ARTIKUL].astype(str).str.lower().str.contains(query, na=False) |
                self.original_df[StandardColumns.NOMENCLATURA].astype(str).str.lower().str.contains(query, na=False) |
                self.original_df.get(StandardColumns.ABC, pd.Series(dtype=str)).astype(str).str.lower().str.contains(
                    query, na=False) |
                self.original_df.get(StandardColumns.XYZ, pd.Series(dtype=str)).astype(str).str.lower().str.contains(
                    query, na=False) |
                self.original_df.get('ABC_XYZ', pd.Series(dtype=str)).astype(str).str.lower().str.contains(query,
                                                                                                           na=False) |
                self.original_df.get('Рекомендация', pd.Series(dtype=str)).astype(str).str.lower().str.contains(query,
                                                                                                                na=False)
        )

        self.df = self.original_df[mask].copy()
        self.update_table()

        # Обновляем счетчик найденных записей
        found_count = len(self.df)
        total_count = len(self.original_df)
        self.search_label.config(text=f'Найдено: {found_count} из {total_count}')

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

    def clear_search(self):
        """Очищает поиск и показывает все данные"""
        self.search_entry.delete(0, tk.END)
        if not self.original_df.empty:
            self.df = self.original_df.copy()
            self.update_table()
            self.search_label.config(text='')

    def try_analyze(self):
        if not self.sales_path or not self.stock_path:
            return

        try:
            sales = load_sales_detailed(self.sales_path)
            stock = load_stock(self.stock_path)

            # ABC-анализ (без изменений)
            abc_df = sales.groupby([
                StandardColumns.ARTIKUL,
                StandardColumns.NOMENCLATURA
            ])[StandardColumns.SUMMA].sum().reset_index()
            abc_df = self.abc_analyzer.analyze(abc_df)

            # XYZ-анализ (ИСПРАВЛЕНО!)
            monthly_sales = sales.groupby([
                StandardColumns.ARTIKUL,
                StandardColumns.NOMENCLATURA,
                "Месяц"
            ])[StandardColumns.SUMMA].sum().reset_index()

            stats = monthly_sales.groupby([
                StandardColumns.ARTIKUL,
                StandardColumns.NOMENCLATURA
            ])[StandardColumns.SUMMA].agg(['mean', 'std', 'count']).reset_index()

            xyz_df = self.xyz_analyzer.analyze(stats)

            # Объединяем результаты
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

        except Exception as e:
            messagebox.showerror('Ошибка анализа', str(e))
            return

        df['ABC_XYZ'] = df[StandardColumns.ABC] + df[StandardColumns.XYZ]

        def get_recommendation(code):
            match code:
                case 'AX':
                    return 'Держим как зеницу ока. Продаётся отлично — пусть лежит.'
                case 'AY' | 'AZ':
                    return 'Хитрый парень: спрос нестабильный. Запас — по ситуации.'
                case 'BX':
                    return 'Норм, но не шик. Присматривай.'
                case 'BY' | 'BZ':
                    return 'Ни рыба, ни мясо. Понаблюдай.'
                case 'CX':
                    return 'Эй, ты чего тут делаешь? Убираем из ассортимента.'
                case 'CY' | 'CZ':
                    return 'Серьёзно? Это ещё живо? Акции, уценка, распродажа!'
                case _:
                    return '¯\_(ツ)_/¯ Нужна экспертная оценка.'

        df['Рекомендация'] = df['ABC_XYZ'].apply(get_recommendation)

        # Сохраняем оригинальные данные для поиска
        self.original_df = df.copy()
        self.df = df
        self.update_table()
        self.status_label.config(text=f'✅ Готово: {len(df)} позиций')

        def get_recommendation(code):
            match code:
                case 'AX':
                    return 'Держим как зеницу ока. Продаётся отлично — пусть лежит.'
                case 'AY' | 'AZ':
                    return 'Хитрый парень: спрос нестабильный. Запас — по ситуации.'
                case 'BX':
                    return 'Норм, но не шик. Присматривай.'
                case 'BY' | 'BZ':
                    return 'Ни рыба, ни мясо. Понаблюдай.'
                case 'CX':
                    return 'Эй, ты чего тут делаешь? Убираем из ассортимента.'
                case 'CY' | 'CZ':
                    return 'Серьёзно? Это ещё живо? Акции, уценка, распродажа!'
                case _:
                    return '¯\_(ツ)_/¯ Нужна экспертная оценка.'

        df['Рекомендация'] = df['ABC_XYZ'].apply(get_recommendation)
        self.df = df
        self.update_table()
        self.status_label.config(text=f'✅ Готово: {len(df)} позиций')

    def update_table(self):
        """Обновляет таблицу и очищает выделение"""
        # Очищаем выделение при обновлении таблицы
        self.clear_cell_selection()
        self.selected_item = None
        self.selected_column = None

        self.tree.delete(*self.tree.get_children())
        for _, row in self.df.iterrows():
            values = [
                row[StandardColumns.ARTIKUL],
                row[StandardColumns.NOMENCLATURA],
                f"{row[StandardColumns.SUMMA]:,.0f}".replace(',', ' '),  # Форматирование с пробелами
                row[StandardColumns.ABC],
                row[StandardColumns.XYZ],
                f"{row[StandardColumns.OSTATOK]:,.0f}".replace(',', ' '),  # Форматирование с пробелами
                row.get('ABC_XYZ', ''),
                row.get('Рекомендация', '')
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

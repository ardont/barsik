# -*- coding: utf-8 -*-
"""
Главное окно графического интерфейса приложения
"""

import os
import re
import json
import threading
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk

from config import WINDOW_TITLE, WINDOW_WIDTH, WINDOW_HEIGHT, MANUAL_LINKS_FILE
from settings_manager import load_settings, save_settings
from engine.loader import load_data
from engine.matcher import match_records
from engine.calculator import calculate_reconciliation
from reports.excel_export import export_to_excel
from reports.word_export import export_to_word
from gui.widgets import KPICard, DetailsPanel

# Настройка внешнего вида customtkinter
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class ReconciliationApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title(WINDOW_TITLE)
        self.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        
        # Загружаем настройки папки по умолчанию
        self.default_folder = load_settings()
        
        # Переменные путей
        self.input_file = ctk.StringVar()
        self.output_excel = ctk.StringVar()
        self.output_word = ctk.StringVar()
        
        # Кэш ручных связей
        self.manual_links = self.load_manual_links()
        
        # Данные сверки
        self.tp_items = []
        self.bt_items = []
        self.matches = []
        self.unmatched_tp = []
        self.unmatched_bt = []
        self.summary = None
        
        self.setup_ui()
        
    def load_manual_links(self) -> dict:
        if os.path.exists(MANUAL_LINKS_FILE):
            try:
                with open(MANUAL_LINKS_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                pass
        return {}
        
    def save_manual_links(self) -> None:
        try:
            with open(MANUAL_LINKS_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.manual_links, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить ручные связи:\n{e}")
            
    def setup_ui(self):
        # Конфигурируем сетку главного окна
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        # ----------------------------------------------------
        # 1. Верхняя панель управления (Выбор файлов)
        # ----------------------------------------------------
        self.ctrl_frame = ctk.CTkFrame(self)
        self.ctrl_frame.grid(row=0, column=0, sticky="ew", padx=15, pady=(15, 5))
        self.ctrl_frame.grid_columnconfigure(1, weight=1)
        
        # Выбор исходного файла
        lbl_file = ctk.CTkLabel(self.ctrl_frame, text="Исходный реестр Excel:")
        lbl_file.grid(row=0, column=0, padx=(15, 5), pady=10, sticky="w")
        
        self.ent_file = ctk.CTkEntry(self.ctrl_frame, textvariable=self.input_file)
        self.ent_file.grid(row=0, column=1, padx=5, pady=10, sticky="ew")
        
        btn_browse = ctk.CTkButton(self.ctrl_frame, text="Обзор...", command=self.browse_input, width=90)
        btn_browse.grid(row=0, column=2, padx=(5, 15), pady=10)
        
        # Кнопки действий
        self.btn_run = ctk.CTkButton(
            self.ctrl_frame, 
            text="▶ Сверить данные", 
            fg_color="#2E7D32", 
            hover_color="#1B5E20",
            command=self.start_analysis
        )
        self.btn_run.grid(row=0, column=3, padx=5, pady=10)
        
        self.btn_export_xls = ctk.CTkButton(
            self.ctrl_frame, 
            text="📊 Экспорт Excel", 
            state="disabled",
            command=self.export_excel
        )
        self.btn_export_xls.grid(row=0, column=4, padx=5, pady=10)
        
        self.btn_export_doc = ctk.CTkButton(
            self.ctrl_frame, 
            text="📋 Экспорт Word", 
            state="disabled",
            command=self.export_word
        )
        self.btn_export_doc.grid(row=0, column=5, padx=(5, 15), pady=10)
        
        # Переключатель темы
        self.btn_theme = ctk.CTkButton(self.ctrl_frame, text="🌓 Тема", width=60, command=self.toggle_theme)
        self.btn_theme.grid(row=0, column=6, padx=(0, 15), pady=10)
        
        # ----------------------------------------------------
        # 2. Панель KPI показателей
        # ----------------------------------------------------
        self.kpi_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.kpi_frame.grid(row=1, column=0, sticky="ew", padx=15, pady=5)
        for c in range(4):
            self.kpi_frame.grid_columnconfigure(c, weight=1)
            
        self.kpi_total = KPICard(self.kpi_frame, "ВСЕГО ПОЗИЦИЙ", "0 (0 TP / 0 BT)")
        self.kpi_total.grid(row=0, column=0, padx=(0, 5), sticky="ew")
        
        self.kpi_rate = KPICard(self.kpi_frame, "ПРОЦЕНТ СОВПАДЕНИЙ", "0.0%")
        self.kpi_rate.grid(row=0, column=1, padx=5, sticky="ew")
        
        self.kpi_profit = KPICard(self.kpi_frame, "ИТОГОВАЯ ПРИБЫЛЬ", "0.00 руб.", text_color="#A5D6A7")
        self.kpi_profit.grid(row=0, column=2, padx=5, sticky="ew")
        
        self.kpi_discrepancy = KPICard(self.kpi_frame, "СУММА РАСХОЖДЕНИЙ", "0.00 руб.", text_color="#EF9A9A")
        self.kpi_discrepancy.grid(row=0, column=3, padx=(5, 0), sticky="ew")
        
        # ----------------------------------------------------
        # 3. Вкладки результатов (Notebook)
        # ----------------------------------------------------
        self.tab_view = ctk.CTkTabview(self)
        self.tab_view.grid(row=2, column=0, sticky="nsew", padx=15, pady=5)
        
        tab_all = self.tab_view.add("📊 Все сопоставления")
        tab_un_tp = self.tab_view.add("❌ Только в TicketProf")
        tab_un_bt = self.tab_view.add("❌ Только в Bars Tour")
        tab_links = self.tab_view.add("⛓ Ручное сопоставление")
        
        # Настройка сеток во вкладках
        for tab in [tab_all, tab_un_tp, tab_un_bt, tab_links]:
            tab.grid_rowconfigure(0, weight=1)
            tab.grid_columnconfigure(0, weight=1)
            
        # Стилизуем стандартный Treeview для темной темы
        self.style_treeviews()
        
        # Вкладка 1: Главная таблица сопоставлений
        self.tree_all = self.create_treeview(tab_all, [
            ("ID", 100), ("Type", 80), 
            ("TP Desc", 260), ("TP Amt", 110), 
            ("BT Desc", 260), ("BT Amt", 110), 
            ("Profit", 100), ("Method", 120), ("Status", 120)
        ])
        self.tree_all.bind("<<TreeviewSelect>>", lambda e: self.on_row_select(self.tree_all))
        
        # Вкладка 2: Нераспределенные TicketProf
        self.tree_tp = self.create_treeview(tab_un_tp, [
            ("Row", 60), ("Date", 90), ("Doc", 180),
            ("Nomenclature", 450), ("Type", 100), ("Amount", 120), ("ID", 120)
        ])
        self.tree_tp.bind("<<TreeviewSelect>>", lambda e: self.on_row_select(self.tree_tp))
        
        # Вкладка 3: Нераспределенные Bars Tour
        self.tree_bt = self.create_treeview(tab_un_bt, [
            ("Row", 60), ("Date", 90), ("Doc", 180),
            ("Nomenclature", 450), ("Type", 100), ("Amount", 120), ("ID", 120)
        ])
        self.tree_bt.bind("<<TreeviewSelect>>", lambda e: self.on_row_select(self.tree_bt))
        
        # Вкладка 4: Ручной сопоставитель (Сплит на две таблицы)
        self.setup_manual_links_tab(tab_links)
        
        # ----------------------------------------------------
        # 4. Панель детальной информации
        # ----------------------------------------------------
        self.details_panel = DetailsPanel(self)
        self.details_panel.grid(row=3, column=0, sticky="ew", padx=15, pady=(5, 10))
        
        # ----------------------------------------------------
        # 5. Статус бар и прогресс
        # ----------------------------------------------------
        self.status_frame = ctk.CTkFrame(self, height=20, fg_color="transparent")
        self.status_frame.grid(row=4, column=0, sticky="ew", padx=15, pady=(0, 10))
        self.status_frame.grid_columnconfigure(0, weight=1)
        
        self.lbl_status = ctk.CTkLabel(self.status_frame, text="Готов к работе", font=ctk.CTkFont(size=10))
        self.lbl_status.grid(row=0, column=0, sticky="w")
        
        self.progress_bar = ctk.CTkProgressBar(self.status_frame, height=8, width=200)
        self.progress_bar.grid(row=0, column=1, sticky="e")
        self.progress_bar.set(0)
        
    def style_treeviews(self):
        style = ttk.Style()
        style.theme_use("clam")
        
        # Стилизация таблицы для темной темы
        style.configure("Treeview",
            background="#1E1E1E",
            foreground="#D4D4D4",
            fieldbackground="#1E1E1E",
            font=("Arial", 9),
            rowheight=24
        )
        style.configure("Treeview.Heading",
            background="#2D2D2D",
            foreground="#FFFFFF",
            font=("Arial", 9, "bold"),
            borderwidth=1
        )
        style.map("Treeview.Heading", background=[('active', '#3E3E3E')])
        style.map("Treeview", background=[('selected', '#0A5F9E')], foreground=[('selected', '#FFFFFF')])
        
    def create_treeview(self, parent, columns_info) -> ttk.Treeview:
        cols = [info[0] for info in columns_info]
        tree = ttk.Treeview(parent, columns=cols, show="headings", selectmode="browse")
        tree.grid(row=0, column=0, sticky="nsew")
        
        # Добавляем скроллбары
        vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        hsb = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        hsb.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        for col_name, width in columns_info:
            tree.heading(col_name, text=col_name)
            tree.column(col_name, width=width, anchor="center" if col_name in ["ID", "Row", "Date", "Type", "Status", "Method"] else "w")
            
        return tree
        
    def setup_manual_links_tab(self, tab):
        tab.grid_columnconfigure(0, weight=4)
        tab.grid_columnconfigure(1, weight=1)
        tab.grid_columnconfigure(2, weight=4)
        tab.grid_rowconfigure(0, weight=1)
        
        # Левая таблица: Несопоставленный TicketProf
        lf = ctk.CTkFrame(tab)
        lf.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        lf.grid_rowconfigure(1, weight=1)
        lf.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(lf, text="Нераспределенный TicketProf", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, pady=5)
        self.tree_manual_tp = self.create_treeview(lf, [("Row", 50), ("Nomenclature", 300), ("Amount", 90)])
        
        # Правая таблица: Несопоставленный Bars Tour
        rf = ctk.CTkFrame(tab)
        rf.grid(row=0, column=2, sticky="nsew", padx=(5, 0))
        rf.grid_rowconfigure(1, weight=1)
        rf.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(rf, text="Нераспределенный Bars Tour", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, pady=5)
        self.tree_manual_bt = self.create_treeview(rf, [("Row", 50), ("Nomenclature", 300), ("Amount", 90)])
        
        # Центральная колонка с кнопкой "Связать" и списком связей
        center_frame = ctk.CTkFrame(tab)
        center_frame.grid(row=0, column=1, sticky="nsew", padx=5)
        center_frame.grid_columnconfigure(0, weight=1)
        center_frame.grid_rowconfigure(2, weight=1)
        
        btn_link = ctk.CTkButton(
            center_frame, 
            text="⛓ Связать", 
            fg_color="#0D47A1", 
            hover_color="#1565C0",
            command=self.create_manual_link
        )
        btn_link.grid(row=0, column=0, pady=20, padx=10, sticky="ew")
        
        ctk.CTkLabel(center_frame, text="Активные ручные связи:", font=ctk.CTkFont(size=11, weight="bold")).grid(row=1, column=0, pady=(10, 2))
        
        self.lst_links = ctk.CTkTextbox(center_frame, font=ctk.CTkFont(size=10), fg_color="#1E1E1E")
        self.lst_links.grid(row=2, column=0, pady=5, padx=10, sticky="nsew")
        self.lst_links.configure(state="disabled")
        
        btn_clear_link = ctk.CTkButton(
            center_frame, 
            text="❌ Удалить связь", 
            fg_color="#B71C1C", 
            hover_color="#C62828",
            command=self.delete_selected_link
        )
        btn_clear_link.grid(row=3, column=0, pady=10, padx=10, sticky="ew")
        
    def browse_input(self):
        file = filedialog.askopenfilename(
            initialdir=self.default_folder,
            title="Выберите исходный Excel-файл",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if file:
            self.input_file.set(file)
            # Запоминаем папку
            self.default_folder = os.path.dirname(file)
            save_settings(self.default_folder)
            
            # Предлагаем дефолтные пути сохранения
            base = os.path.splitext(file)[0]
            self.output_excel.set(f"{base}_сопоставлено.xlsx")
            self.output_word.set(f"{base}_акт_сверки.docx")
            
    def toggle_theme(self):
        current = ctk.get_appearance_mode()
        new_mode = "Light" if current == "Dark" else "Dark"
        ctk.set_appearance_mode(new_mode)
        
    def start_analysis(self):
        file_path = self.input_file.get()
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("Ошибка", "Пожалуйста, выберите существующий файл Excel.")
            return
            
        self.btn_run.configure(state="disabled")
        self.lbl_status.configure(text="Чтение файла Excel...")
        self.progress_bar.set(0.2)
        
        # Выполняем в отдельном потоке
        threading.Thread(target=self.run_reconciliation_thread, args=(file_path,), daemon=True).start()
        
    def run_reconciliation_thread(self, file_path: str):
        try:
            # 1. Загрузка
            tp_items, bt_items = load_data(file_path)
            self.tp_items = tp_items
            self.bt_items = bt_items
            
            self.lbl_status.configure(text="Вычисление соответствий...")
            self.progress_bar.set(0.5)
            
            # 2. Сопоставление
            self.matches, self.unmatched_tp, self.unmatched_bt = match_records(
                tp_items, bt_items, self.manual_links
            )
            
            # 3. Финансовый расчет
            self.summary = calculate_reconciliation(self.tp_items, self.bt_items, self.matches)
            
            # Обновление графического интерфейса
            self.after(0, self.on_analysis_complete)
        except Exception as e:
            self.after(0, self.on_analysis_error, str(e))
            
    def on_analysis_complete(self):
        self.btn_run.configure(state="normal")
        self.btn_export_xls.configure(state="normal")
        self.btn_export_doc.configure(state="normal")
        
        self.lbl_status.configure(text="Сверка успешно завершена!")
        self.progress_bar.set(1.0)
        
        # Обновляем KPI
        total_str = f"{len(self.tp_items) + len(self.bt_items)} ({len(self.tp_items)} TP / {len(self.bt_items)} BT)"
        self.kpi_total.update_value(total_str)
        
        # Вычисляем процент совпадений для TicketProf
        rate_val = (self.summary.matched_tp_count / self.summary.total_tp_count * 100) if self.summary.total_tp_count else 0.0
        self.kpi_rate.update_value(f"{rate_val:.1f}%")
        
        self.kpi_profit.update_value(f"{self.summary.total_profit:,.2f} руб.")
        self.kpi_discrepancy.update_value(f"{self.summary.unmatched_tp_sum + self.summary.unmatched_bt_sum:,.2f} руб.")
        
        # Заполняем таблицы
        self.populate_grids()
        
    def populate_grids(self):
        # 1. Сбрасываем старые значения
        for tree in [self.tree_all, self.tree_tp, self.tree_bt, self.tree_manual_tp, self.tree_manual_bt]:
            for row in tree.get_children():
                tree.delete(row)
                
        # 2. Вкладка "Все сопоставления"
        # Настройка тегов строк для раскраски
        self.tree_all.tag_configure("matched", background="#1E3E20", foreground="#A5D6A7")
        self.tree_all.tag_configure("unmatched", background="#3E2723", foreground="#FFAB91")
        self.tree_all.tag_configure("discrepancy", background="#4A148C", foreground="#E1BEE7")
        
        for tp, bt, method, score in self.matches:
            # Для отелей проверяем аномалию
            is_anomaly = False
            if tp.service_type == "Hotel":
                expected_profit = 0.1 * bt.amount
                if abs((bt.amount - tp.allocated_amount) - expected_profit) > 0.01:
                    is_anomaly = True
                    
            tag = "discrepancy" if is_anomaly else "matched"
            tp_id = list(tp.ids)[0] if tp.ids else "N/A"
            self.tree_all.insert("", "end", values=(
                tp_id, tp.service_type,
                tp.desc, f"{tp.allocated_amount:,.2f}",
                bt.desc, f"{bt.amount:,.2f}",
                f"{(bt.amount - tp.allocated_amount):,.2f}",
                method, "Сверка успешна" if not is_anomaly else "Проверьте маржу"
            ), tags=(tag,))
            
        # Добавляем нераспределенные в общую таблицу
        for tp in self.unmatched_tp:
            tp_id = list(tp.ids)[0] if tp.ids else "N/A"
            self.tree_all.insert("", "end", values=(
                tp_id, tp.service_type,
                tp.desc, f"{tp.allocated_amount:,.2f}",
                "Отсутствует в Bars Tour", "0.00", "0.00",
                "Не сопоставлено", "Только в TicketProf"
            ), tags=("unmatched",))
            
        for bt in self.unmatched_bt:
            bt_id = list(bt.ids)[0] if bt.ids else "N/A"
            self.tree_all.insert("", "end", values=(
                bt_id, bt.service_type,
                "Отсутствует в TicketProf", "0.00",
                bt.desc, f"{bt.amount:,.2f}", "0.00",
                "Не сопоставлено", "Только в Bars Tour"
            ), tags=("unmatched",))
            
        # 3. Вкладки "Только TicketProf" и "Только Bars Tour"
        for tp in self.unmatched_tp:
            tp_id = list(tp.ids)[0] if tp.ids else "N/A"
            self.tree_tp.insert("", "end", values=(
                tp.row, tp.date, tp.doc, tp.desc, tp.service_type, f"{tp.allocated_amount:,.2f}", tp_id
            ))
            # Заполняем также левую часть ручного сопоставителя
            self.tree_manual_tp.insert("", "end", values=(tp.row, tp.desc, f"{tp.allocated_amount:,.2f}"))
            
        for bt in self.unmatched_bt:
            bt_id = list(bt.ids)[0] if bt.ids else "N/A"
            self.tree_bt.insert("", "end", values=(
                bt.row, bt.date, bt.doc, bt.desc, bt.service_type, f"{bt.amount:,.2f}", bt_id
            ))
            # Заполняем также правую часть ручного сопоставителя
            self.tree_manual_bt.insert("", "end", values=(bt.row, bt.desc, f"{bt.amount:,.2f}"))
            
        # Обновляем текстовый блок активных связей
        self.update_manual_links_list()
        
    def update_manual_links_list(self):
        self.lst_links.configure(state="normal")
        self.lst_links.delete("1.0", "end")
        
        if not self.manual_links:
            self.lst_links.insert("end", "Нет сохраненных ручных связей.\n")
        else:
            for tp_clean, bt_clean in self.manual_links.items():
                self.lst_links.insert("end", f"TP: «{tp_clean[:30]}...»\n ➔ BT: «{bt_clean[:30]}...»\n\n")
                
        self.lst_links.configure(state="disabled")
        
    def on_analysis_error(self, error_str: str):
        self.btn_run.configure(state="normal")
        self.lbl_status.configure(text="Произошла ошибка при анализе")
        self.progress_bar.set(0)
        messagebox.showerror("Ошибка сверки", f"Во время анализа возникла критическая ошибка:\n{error_str}")
        
    def on_row_select(self, tree: ttk.Treeview):
        selected = tree.selection()
        if not selected:
            self.details_panel.clear()
            return
            
        item = tree.item(selected[0])
        values = item["values"]
        
        # Составляем детальный текст
        details = []
        if len(values) == 9: # Из таблицы "Все сопоставления"
            details.append(f"Код идентификации: {values[0]}")
            details.append(f"Категория услуги: {values[1]}")
            details.append(f"--- TicketProf ---")
            details.append(f"Описание: {values[2]}")
            details.append(f"Стоимость (нетто): {values[3]}")
            details.append(f"--- Bars Tour ---")
            details.append(f"Описание: {values[4]}")
            details.append(f"Сумма (брутто): {values[5]}")
            details.append(f"Прибыль: {values[6]} руб.")
            details.append(f"Метод привязки: {values[7]}")
            details.append(f"Статус: {values[8]}")
        else: # Из таблиц расхождений
            details.append(f"Строка в Excel: {values[0]}")
            details.append(f"Дата транзакции: {values[1]}")
            details.append(f"Документ: {values[2]}")
            details.append(f"Номенклатура (описание): {values[3]}")
            details.append(f"Тип услуги: {values[4]}")
            details.append(f"Сумма: {values[5]} руб.")
            details.append(f"ID: {values[6]}")
            
        self.details_panel.show_details("\n".join(details))
        
    def create_manual_link(self):
        sel_tp = self.tree_manual_tp.selection()
        sel_bt = self.tree_manual_bt.selection()
        
        if not sel_tp or not sel_bt:
            messagebox.showwarning("Внимание", "Пожалуйста, выберите по одной записи в левой и правой таблицах для связывания.")
            return
            
        tp_vals = self.tree_manual_tp.item(sel_tp[0])["values"]
        bt_vals = self.tree_manual_bt.item(sel_bt[0])["values"]
        
        # Находим исходные ServiceItem
        tp_row = int(tp_vals[0])
        bt_row = int(bt_vals[0])
        
        tp_item = next(item for item in self.tp_items if item.row == tp_row)
        bt_item = next(item for item in self.bt_items if item.row == bt_row)
        
        # Добавляем в словарь
        self.manual_links[tp_item.clean_desc] = bt_item.clean_desc
        self.save_manual_links()
        
        # Перезапускаем сверку
        self.start_analysis()
        messagebox.showinfo("Готово", f"Связь успешно создана!\nПересчет завершен.")
        
    def delete_selected_link(self):
        if not self.manual_links:
            return
            
        # Запрашиваем подтверждение
        confirm = messagebox.askyesno("Удаление ручных связей", "Вы действительно хотите удалить ВСЕ ручные связи и вернуть исходное состояние?")
        if confirm:
            self.manual_links.clear()
            self.save_manual_links()
            self.start_analysis()
            messagebox.showinfo("Готово", "Все ручные связи очищены.")
            
    def export_excel(self):
        out_path = self.output_excel.get()
        if not out_path:
            return
            
        try:
            export_to_excel(
                self.tp_items, self.bt_items, self.matches, 
                self.unmatched_tp, self.unmatched_bt, self.summary, 
                out_path
            )
            messagebox.showinfo("Готово", f"Отчет Excel успешно сохранен:\n{out_path}")
        except Exception as e:
            messagebox.showerror("Ошибка сохранения", f"Не удалось экспортировать отчет:\n{e}")
            
    def export_word(self):
        out_path = self.output_word.get()
        if not out_path:
            return
            
        try:
            export_to_word(
                self.unmatched_tp, self.unmatched_bt, self.summary, 
                out_path
            )
            messagebox.showinfo("Готово", f"Официальный Акт сверки успешно сохранен:\n{out_path}")
        except Exception as e:
            messagebox.showerror("Ошибка сохранения", f"Не удалось создать Акт сверки:\n{e}")

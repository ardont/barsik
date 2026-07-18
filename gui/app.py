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
import tkinter as tk
from PIL import Image

from config import WINDOW_TITLE, WINDOW_WIDTH, WINDOW_HEIGHT, MANUAL_LINKS_FILE
from settings_manager import load_settings, save_settings
from engine.loader import load_data
from engine.matcher import match_records
from engine.calculator import calculate_reconciliation
from reports.excel_export import export_to_excel
from reports.word_export import export_to_word
from gui.widgets import KPICard, DetailsPanel

# Настройка внешнего вида customtkinter
ctk.set_default_color_theme("blue")

class ReconciliationApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title(WINDOW_TITLE)
        self.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        
        # Устанавливаем иконку приложения если она есть
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets", "barsik_logo.png")
        if os.path.exists(icon_path):
            try:
                self.iconphoto(False, tk.PhotoImage(file=icon_path))
            except Exception:
                pass
                
        # Загружаем настройки
        self.settings = load_settings()
        self.default_folder = self.settings['default_folder']
        from models import ServiceItem
        ServiceItem.hotel_margin = self.settings.get('hotel_margin', 10.0)
        
        # Устанавливаем сохраненную тему оформления
        ctk.set_appearance_mode(self.settings.get('theme', 'Dark'))
        
        # Переменные путей файлов
        self.input_file_tp = ctk.StringVar()
        self.input_file_bt = ctk.StringVar()
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
            
    def get_asset_image(self, filename: str, size: tuple) -> ctk.CTkImage:
        """
        Безопасная загрузка изображения из папки assets
        """
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets", filename)
        if os.path.exists(path):
            try:
                img = Image.open(path)
                return ctk.CTkImage(light_image=img, dark_image=img, size=size)
            except Exception:
                pass
        return None
        
    def setup_ui(self):
        # Конфигурируем сетку главного окна
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        # ----------------------------------------------------
        # 1. Вверхняя панель управления (Выбор файлов)
        # ----------------------------------------------------
        self.ctrl_frame = ctk.CTkFrame(self)
        self.ctrl_frame.grid(row=0, column=0, sticky="ew", padx=15, pady=(15, 5))
        self.ctrl_frame.grid_columnconfigure(1, weight=1)
        
        # Отображение логотипа Барсика в левом верхнем углу
        logo_img = self.get_asset_image("barsik_logo.png", (48, 48))
        if logo_img:
            self.lbl_logo = ctk.CTkLabel(self.ctrl_frame, image=logo_img, text="")
            self.lbl_logo.grid(row=0, column=0, padx=(15, 5), pady=15)
            col_offset = 1
        else:
            col_offset = 0
            
        # Строка 1: Сводный файл сверки (Excel)
        lbl_file_tp = ctk.CTkLabel(self.ctrl_frame, text="Сводный файл сверки (Excel):")
        lbl_file_tp.grid(row=0, column=col_offset, padx=(15, 5), pady=15, sticky="w")
        
        self.ent_file_tp = ctk.CTkEntry(self.ctrl_frame, textvariable=self.input_file_tp, placeholder_text="Выберите единый файл сверки (например, 08.07_1.xlsx)...")
        self.ent_file_tp.grid(row=0, column=col_offset + 1, padx=5, pady=15, sticky="ew")
        
        btn_browse_tp = ctk.CTkButton(self.ctrl_frame, text="Обзор...", command=self.browse_tp, width=90)
        btn_browse_tp.grid(row=0, column=col_offset + 2, padx=(5, 5), pady=15)
        
        self.lbl_tp_check = ctk.CTkLabel(self.ctrl_frame, text="", text_color="#A5D6A7")
        self.lbl_tp_check.grid(row=0, column=col_offset + 3, padx=(5, 15), pady=15, sticky="w")
        
        # Действия и Кнопки
        self.actions_frame = ctk.CTkFrame(self.ctrl_frame, fg_color="transparent")
        self.actions_frame.grid(row=0, column=col_offset + 4, padx=(15, 15), pady=10, sticky="ns")
        
        self.btn_run = ctk.CTkButton(
            self.actions_frame, 
            text="▶ Запустить анализ", 
            fg_color="#2E7D32", 
            hover_color="#1B5E20",
            command=self.start_analysis,
            width=150
        )
        self.btn_run.pack(side="top", pady=2)
        
        self.btn_export_xls = ctk.CTkButton(
            self.actions_frame, 
            text="📊 Экспорт Excel", 
            state="disabled",
            command=self.export_excel,
            width=150
        )
        self.btn_export_xls.pack(side="top", pady=2)
        
        self.btn_export_doc = ctk.CTkButton(
            self.actions_frame, 
            text="📋 Экспорт Word", 
            state="disabled",
            command=self.export_word,
            width=150
        )
        self.btn_export_doc.pack(side="top", pady=2)
        
        # ----------------------------------------------------
        # 2. Панель KPI показателей
        # ----------------------------------------------------
        self.kpi_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.kpi_frame.grid(row=1, column=0, sticky="ew", padx=15, pady=5)
        for c in range(4):
            self.kpi_frame.grid_columnconfigure(c, weight=1)
            
        self.kpi_total = KPICard(self.kpi_frame, "ВСЕГО ПОЗИЦИЙ", "0 (0 TP / 0 BT)")
        self.kpi_total.grid(row=0, column=0, padx=(0, 5), sticky="ew")
        
        self.kpi_rate = KPICard(self.kpi_frame, "СТОИМОСТЬ УСЛУГ (TP)", "0.00 руб.")
        self.kpi_rate.grid(row=0, column=1, padx=5, sticky="ew")
        
        self.kpi_profit = KPICard(self.kpi_frame, "ИТОГО В БАРСЕ (BT)", "0.00 руб.", text_color=("#0A5F9E", "#90CAF9"))
        self.kpi_profit.grid(row=0, column=2, padx=5, sticky="ew")
        
        self.kpi_discrepancy = KPICard(self.kpi_frame, "ПРИБЫЛЬ", "0.00 руб.", text_color=("#2E7D32", "#A5D6A7"))
        self.kpi_discrepancy.grid(row=0, column=3, padx=(5, 0), sticky="ew")
        
        # ----------------------------------------------------
        # 3. Вкладки результатов (Notebook)
        # ----------------------------------------------------
        self.tab_view = ctk.CTkTabview(self)
        self.tab_view.grid(row=2, column=0, sticky="nsew", padx=15, pady=5)
        
        self.tab_all = self.tab_view.add("📊 Все сопоставления")
        self.tab_mismatches = self.tab_view.add("⚠ Несоответствия")
        tab_un_tp = self.tab_view.add("🟡 В Тикете, нет в Барсе")
        tab_un_bt = self.tab_view.add("🟡 В Барсе, нет в Тикете")
        tab_links = self.tab_view.add("⛓ Ручное сопоставление")
        tab_help = self.tab_view.add("❓ Справка")
        tab_settings = self.tab_view.add("🔧 Настройки сверки")
        
        # Настройка сеток во вкладках
        for tab in [self.tab_all, self.tab_mismatches, tab_un_tp, tab_un_bt, tab_links, tab_help, tab_settings]:
            tab.grid_rowconfigure(0, weight=1)
            tab.grid_columnconfigure(0, weight=1)
            
        self.style_treeviews()
        
        # Вкладка 1: Главная таблица сопоставлений
        self.tree_all = self.create_treeview(self.tab_all, [
            ("ID", 100), ("Тип услуги", 85), 
            ("Документ TicketProf", 130), ("Номенклатура TicketProf", 230), ("Стоимость услуг", 110), 
            ("Документ Bars Tour", 130), ("Номенклатура Bars Tour", 230), ("Итого в Барсе", 110), 
            ("Прибыль", 100), ("Метод привязки", 130), ("Статус", 150)
        ])
        self.tree_all.bind("<<TreeviewSelect>>", lambda e: self.on_row_select(self.tree_all))
        
        # Вкладка Несоответствия
        self.tree_mismatches = self.create_treeview(self.tab_mismatches, [
            ("ID", 100), ("Тип услуги", 85), 
            ("Документ TicketProf", 130), ("Номенклатура TicketProf", 230), ("Стоимость услуг", 110), 
            ("Документ Bars Tour", 130), ("Номенклатура Bars Tour", 230), ("Итого в Барсе", 110), 
            ("Прибыль", 100), ("Метод привязки", 130), ("Статус", 150)
        ])
        self.tree_mismatches.bind("<<TreeviewSelect>>", lambda e: self.on_row_select(self.tree_mismatches))
        
        # Инициализируем пустое состояние (Empty State) в центре таблицы "Все сопоставления"
        self.setup_empty_state()
        
        # Вкладка 2: В Тикете, нет в Барсе
        self.tree_tp = self.create_treeview(tab_un_tp, [
            ("Строка", 60), ("Дата", 90), ("Документ", 180),
            ("Номенклатура TicketProf", 450), ("Тип услуги", 100), ("Стоимость услуг", 120), ("ID", 120)
        ])
        self.tree_tp.bind("<<TreeviewSelect>>", lambda e: self.on_row_select(self.tree_tp))
        
        # Вкладка 3: В Барсе, нет в Тикете
        self.tree_bt = self.create_treeview(tab_un_bt, [
            ("Строка", 60), ("Дата", 90), ("Документ", 180),
            ("Номенклатура Bars Tour", 450), ("Тип услуги", 100), ("Итого в Барсе", 120), ("ID", 120)
        ])
        self.tree_bt.bind("<<TreeviewSelect>>", lambda e: self.on_row_select(self.tree_bt))
        
        # Вкладка 4: Ручной сопоставитель (Сплит на две таблицы)
        self.setup_manual_links_tab(tab_links)
        
        # Вкладка 5: Справка (Руководство пользователя)
        self.setup_help_tab(tab_help)
        
        # Вкладка 6: Настройки сверки
        self.setup_settings_tab(tab_settings)
        
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
        
        self.lbl_status = ctk.CTkLabel(self.status_frame, text="Ожидание загрузки файлов...", font=ctk.CTkFont(size=10))
        self.lbl_status.grid(row=0, column=0, sticky="w")
        
        self.progress_bar = ctk.CTkProgressBar(self.status_frame, height=8, width=200)
        self.progress_bar.grid(row=0, column=1, sticky="e")
        self.progress_bar.set(0)
        
    def setup_empty_state(self):
        """
        Создает красивое состояние приветствия в центре главной таблицы
        """
        self.empty_state_frame = ctk.CTkFrame(self.tab_all, fg_color=("#F9F9F9", "#1E1E1E"), corner_radius=0)
        self.empty_state_frame.grid(row=0, column=0, sticky="nsew")
        
        self.empty_state_frame.grid_rowconfigure(0, weight=1)
        self.empty_state_frame.grid_rowconfigure(3, weight=1)
        self.empty_state_frame.grid_columnconfigure(0, weight=1)
        
        # Картинка кота-хранителя
        guardian_img = self.get_asset_image("barsik_guardian.png", (220, 220))
        if guardian_img:
            self.lbl_empty_img = ctk.CTkLabel(self.empty_state_frame, image=guardian_img, text="")
            self.lbl_empty_img.grid(row=1, column=0, pady=(50, 10))
            
        self.lbl_empty_text = ctk.CTkLabel(
            self.empty_state_frame, 
            text="Привет! Я Барсик, хранитель финансов туркомпании 🐾\nЗагрузите реестры продаж и приходов, чтобы начать автоматическую сверку!", 
            font=ctk.CTkFont(family="Arial", size=13, weight="bold"),
            text_color=("#555555", "#8A8A8A"),
            justify="center"
        )
        self.lbl_empty_text.grid(row=2, column=0, pady=10)
        
        # Скрываем Treeview до момента первой сверки
        self.tree_all.grid_remove()

    def style_treeviews(self):
        style = ttk.Style()
        style.theme_use("clam")
        
        is_dark = (self.settings.get('theme', 'Dark') == 'Dark')
        
        if is_dark:
            bg_color = "#1E1E1E"
            fg_color = "#D4D4D4"
            heading_bg = "#2D2D2D"
            heading_fg = "#FFFFFF"
            active_heading = "#3E3E3E"
            select_bg = "#0A5F9E"
            select_fg = "#FFFFFF"
        else:
            bg_color = "#FFFFFF"
            fg_color = "#000000"
            heading_bg = "#EAEAEA"
            heading_fg = "#000000"
            active_heading = "#D6D6D6"
            select_bg = "#B0D0F0"
            select_fg = "#000000"
            
        style.configure("Treeview",
            background=bg_color,
            foreground=fg_color,
            fieldbackground=bg_color,
            font=("Arial", 9),
            rowheight=24
        )
        style.configure("Treeview.Heading",
            background=heading_bg,
            foreground=heading_fg,
            font=("Arial", 9, "bold"),
            borderwidth=1
        )
        style.map("Treeview.Heading", background=[('active', active_heading)])
        style.map("Treeview", background=[('selected', select_bg)], foreground=[('selected', select_fg)])
        
    def create_treeview(self, parent, columns_info) -> ttk.Treeview:
        cols = [info[0] for info in columns_info]
        tree = ttk.Treeview(parent, columns=cols, show="headings", selectmode="browse")
        tree.grid(row=0, column=0, sticky="nsew")
        
        vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        hsb = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        hsb.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        for col_name, width in columns_info:
            tree.heading(col_name, text=col_name)
            tree.column(col_name, width=width, anchor="center" if col_name in ["ID", "Строка", "Дата", "Тип услуги", "Статус", "Метод привязки"] else "w")
            
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
        ctk.CTkLabel(lf, text="В Тикете, нет в Барсе", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, pady=5)
        self.tree_manual_tp = self.create_treeview(lf, [("Строка", 50), ("Номенклатура TicketProf", 300), ("Стоимость услуг", 90)])
        
        # Правая таблица: Несопоставленный Bars Tour
        rf = ctk.CTkFrame(tab)
        rf.grid(row=0, column=2, sticky="nsew", padx=(5, 0))
        rf.grid_rowconfigure(1, weight=1)
        rf.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(rf, text="В Барсе, нет в Тикете", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, pady=5)
        self.tree_manual_bt = self.create_treeview(rf, [("Строка", 50), ("Номенклатура Bars Tour", 300), ("Итого в Барсе", 90)])
        
        # Центральная колонка
        center_frame = ctk.CTkFrame(tab)
        center_frame.grid(row=0, column=1, sticky="nsew", padx=5)
        center_frame.grid_columnconfigure(0, weight=1)
        center_frame.grid_rowconfigure(2, weight=1)
        
        btn_link = ctk.CTkButton(
            center_frame, 
            text="⛓ Связать вручную", 
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
            text="❌ Очистить связи", 
            fg_color="#B71C1C", 
            hover_color="#C62828",
            command=self.delete_selected_link
        )
        btn_clear_link.grid(row=3, column=0, pady=10, padx=10, sticky="ew")

    def setup_help_tab(self, tab):
        """
        Создает вкладку интерактивной встроенной справки с разметкой
        """
        tab.grid_columnconfigure(0, weight=3)
        tab.grid_columnconfigure(1, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        
        # Левая часть: Справочная информация
        help_box = tk.Text(
            tab, 
            font=("Arial", 10), 
            insertbackground="#FFFFFF",
            relief="flat", 
            bd=0, 
            highlightthickness=0,
            padx=15,
            pady=15
        )
        self.help_box = help_box
        help_box.grid(row=0, column=0, sticky="nsew", padx=(15, 0), pady=15)
        
        # Добавляем скроллбар
        help_sb = ctk.CTkScrollbar(tab, command=help_box.yview)
        help_sb.grid(row=0, column=0, sticky="nse", padx=(0, 5), pady=15)
        help_box.configure(yscrollcommand=help_sb.set)
        
        # Настройка тегов форматирования
        help_box.tag_config("h1", font=("Arial", 14, "bold"), foreground="#90CAF9", spacing1=10, spacing3=5)
        help_box.tag_config("h2", font=("Arial", 12, "bold"), foreground="#A5D6A7", spacing1=8, spacing3=4)
        help_box.tag_config("body", font=("Arial", 10), foreground="#CCCCCC", spacing1=3, spacing3=3)
        help_box.tag_config("bullet", font=("Arial", 10), foreground="#CCCCCC", lmargin1=20, lmargin2=30, spacing1=2, spacing3=2)
        help_box.tag_config("bold", font=("Arial", 10, "bold"), foreground="#FFFFFF")
        
        # Текст справки
        help_box.insert("end", "📊 Руководство пользователя программы «Умная сверка 3.0»\n\n", "h1")
        
        help_box.insert("end", "1. Описание программы\n", "h2")
        help_box.insert("end", "Программа предназначена для автоматической сверки данных реализации авиабилетов/отелей из системы TicketProf и приходов/себестоимости из Bars Tour. Поиск соответствий идет на уровне номенклатурных позиций с использованием сквозных ID (номера билетов, MCO, заказов гостиниц) и текстового анализа.\n\n", "body")
        
        help_box.insert("end", "2. Глоссарий терминов\n", "h2")
        help_box.insert("end", "• ", "bullet")
        help_box.insert("end", "Стоимость услуг", "bold")
        help_box.insert("end", " — сумма продажи по данным выгрузки TicketProf.\n", "body")
        
        help_box.insert("end", "• ", "bullet")
        help_box.insert("end", "Итого в Барсе", "bold")
        help_box.insert("end", " — сумма прихода или себестоимость услуги по выгрузке Bars Tour.\n", "body")
        
        help_box.insert("end", "• ", "bullet")
        help_box.insert("end", "Прибыль", "bold")
        help_box.insert("end", " — разность (Итого в Барсе - Стоимость услуг) по сопоставленным позициям.\n", "body")
        
        help_box.insert("end", "• ", "bullet")
        help_box.insert("end", "Совпадение", "bold")
        help_box.insert("end", " — услуга успешно привязана по ID или описанию, критических расхождений по суммам нет.\n", "body")
        
        help_box.insert("end", "• ", "bullet")
        help_box.insert("end", "В Тикете, нет в Барсе", "bold")
        help_box.insert("end", " — услуга присутствует в TicketProf, но не найдена в приходе Bars Tour (желтая строка).\n", "body")
        
        help_box.insert("end", "• ", "bullet")
        help_box.insert("end", "В Барсе, нет в Тикете", "bold")
        help_box.insert("end", " — приход по услуге занесен в Bars Tour, но реализация отсутствует (желтая строка).\n", "body")
        
        help_box.insert("end", "• ", "bullet")
        help_box.insert("end", "Несовпадение по суммам", "bold")
        help_box.insert("end", " — услуга сопоставлена, но маржа отрицательна или не сходится с ожиданиями (красная строка).\n\n", "body")
        
        help_box.insert("end", "3. Пошаговая инструкция\n", "h2")
        help_box.insert("end", "Шаг 1. ", "bold")
        help_box.insert("end", "Нажмите кнопку «Обзор...» напротив поля «Сводный файл сверки» и выберите ваш файл Excel.\n", "body")
        help_box.insert("end", "Шаг 2. ", "bold")
        help_box.insert("end", "Нажмите «▶ Запустить анализ». Программа выполнит расчет. Таблицы и KPI-карточки заполнятся автоматически.\n", "body")
        help_box.insert("end", "Шаг 3. ", "bold")
        help_box.insert("end", "Для сопоставления расхождений вручную перейдите на вкладку «Ручное сопоставление», выделите одну позицию слева, одну справа и нажмите «Связать вручную». Программа запомнит это правило и автоматически пересчитает сверку.\n\n", "body")
        
        help_box.insert("end", "4. Пояснения по маржинальности отелей\n", "h2")
        help_box.insert("end", "Для отелей правильная наценка (прибыль) должна составлять ровно 10% от брутто-суммы. Если менеджеры допустили ошибку в формуле или округлении в Excel, строка получит статус «Нетипичная маржа» и окрасится в красный цвет для проверки бухгалтером.\n\n", "body")
        
        help_box.configure(state="disabled")
        self.update_help_box_colors()
        
        # Правая часть: Маскот Барсик-пилот
        pilot_frame = ctk.CTkFrame(tab, fg_color="transparent")
        pilot_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 15), pady=15)
        
        pilot_frame.grid_rowconfigure(0, weight=1)
        pilot_frame.grid_rowconfigure(2, weight=1)
        pilot_frame.grid_columnconfigure(0, weight=1)
        
        pilot_img = self.get_asset_image("barsik_pilot.png", (180, 180))
        if pilot_img:
            self.lbl_pilot_img = ctk.CTkLabel(pilot_frame, image=pilot_img, text="")
            self.lbl_pilot_img.grid(row=0, column=0, sticky="s", pady=10)
            
        self.lbl_pilot_caption = ctk.CTkLabel(
            pilot_frame, 
            text="Барсик-пилот желает вам\nлегкой сверки! ✈️🐾", 
            font=ctk.CTkFont(family="Arial", size=11, slant="italic"),
            text_color="#8A8A8A",
            justify="center"
        )
        self.lbl_pilot_caption.grid(row=1, column=0, sticky="n", pady=5)
        
    def setup_settings_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        
        scroll_frame = ctk.CTkScrollableFrame(tab, fg_color="transparent")
        scroll_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        scroll_frame.grid_columnconfigure(0, weight=1)
        
        # --- Блок 0: Режим работы и Тема ---
        mode_frame = ctk.CTkFrame(scroll_frame)
        mode_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        mode_frame.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(mode_frame, text="⚙️ Общие настройки", font=ctk.CTkFont(size=14, weight="bold"), text_color=("#0D47A1", "#90CAF9")).grid(row=0, column=0, columnspan=2, padx=15, pady=(15, 5), sticky="w")
        
        self.var_simple_mode = tk.BooleanVar(value=self.settings.get('simple_mode', True))
        self.chk_simple_mode = ctk.CTkCheckBox(mode_frame, text="Простой и безопасный режим (включен по умолчанию)", variable=self.var_simple_mode, command=self.toggle_simple_mode_ui)
        self.chk_simple_mode.grid(row=1, column=0, columnspan=2, padx=15, pady=10, sticky="w")
        
        ctk.CTkLabel(mode_frame, text="Цветовая тема оформления:").grid(row=2, column=0, padx=15, pady=10, sticky="w")
        self.var_theme = ctk.StringVar(value=self.settings.get('theme', 'Dark'))
        self.opt_theme = ctk.CTkOptionMenu(mode_frame, values=["Dark", "Light"], variable=self.var_theme, command=self.change_theme)
        self.opt_theme.grid(row=2, column=1, padx=15, pady=10, sticky="w")
        
        # --- Блок 1: Настройки отелей ---
        self.hotel_frame = ctk.CTkFrame(scroll_frame)
        self.hotel_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=10)
        self.hotel_frame.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(self.hotel_frame, text="🏨 Настройки маржи отелей", font=ctk.CTkFont(size=14, weight="bold"), text_color=("#2E7D32", "#A5D6A7")).grid(row=0, column=0, columnspan=2, padx=15, pady=(15, 5), sticky="w")
        
        ctk.CTkLabel(self.hotel_frame, text="Ожидаемая наценка отеля (в %):").grid(row=1, column=0, padx=15, pady=10, sticky="w")
        
        self.var_hotel_margin = tk.DoubleVar(value=self.settings.get('hotel_margin', 10.0))
        self.spn_hotel_margin = ctk.CTkEntry(self.hotel_frame, textvariable=self.var_hotel_margin, width=80)
        self.spn_hotel_margin.grid(row=1, column=1, padx=15, pady=10, sticky="w")
        
        lbl_hotel_hint = ctk.CTkLabel(
            self.hotel_frame, 
            text="Используется для проверки статуса отелей. Если прибыль / стоимость в Барсе\nне совпадает с этим процентом, строка будет помечена красным.", 
            font=ctk.CTkFont(size=11), 
            text_color="#8A8A8A",
            justify="left"
        )
        lbl_hotel_hint.grid(row=2, column=0, columnspan=2, padx=15, pady=(0, 15), sticky="w")
        
        # --- Блок 2: Алгоритмы сопоставления ---
        self.algo_frame = ctk.CTkFrame(scroll_frame)
        self.algo_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=10)
        self.algo_frame.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(self.algo_frame, text="⚙️ Алгоритмы сопоставления", font=ctk.CTkFont(size=14, weight="bold"), text_color=("#0D47A1", "#90CAF9")).grid(row=0, column=0, columnspan=2, padx=15, pady=(15, 5), sticky="w")
        
        self.var_enable_id = tk.BooleanVar(value=self.settings.get('enable_id_match', True))
        self.chk_enable_id = ctk.CTkCheckBox(self.algo_frame, text="Включить сопоставление по уникальным ID (билеты/заказы)", variable=self.var_enable_id)
        self.chk_enable_id.grid(row=1, column=0, columnspan=2, padx=15, pady=5, sticky="w")
        
        self.var_enable_exact = tk.BooleanVar(value=self.settings.get('enable_exact_match', True))
        self.chk_enable_exact = ctk.CTkCheckBox(self.algo_frame, text="Включить точное сопоставление по очищенным названиям", variable=self.var_enable_exact)
        self.chk_enable_exact.grid(row=2, column=0, columnspan=2, padx=15, pady=5, sticky="w")
        
        self.var_enable_fuzzy = tk.BooleanVar(value=self.settings.get('enable_fuzzy_match', True))
        self.chk_enable_fuzzy = ctk.CTkCheckBox(self.algo_frame, text="Включить нечеткое (Fuzzy) сопоставление названий", variable=self.var_enable_fuzzy)
        self.chk_enable_fuzzy.grid(row=3, column=0, columnspan=2, padx=15, pady=5, sticky="w")
        
        ctk.CTkLabel(self.algo_frame, text="Порог схожести текста (в %):").grid(row=4, column=0, padx=15, pady=10, sticky="w")
        
        self.var_fuzzy_threshold = tk.DoubleVar(value=self.settings.get('fuzzy_threshold', 75.0))
        self.spn_fuzzy_threshold = ctk.CTkEntry(self.algo_frame, textvariable=self.var_fuzzy_threshold, width=80)
        self.spn_fuzzy_threshold.grid(row=4, column=1, padx=15, pady=10, sticky="w")
        
        lbl_fuzzy_hint = ctk.CTkLabel(
            self.algo_frame, 
            text="Используется при нечетком поиске. Чем ниже порог, тем больше совпадений будет найдено,\nно выше вероятность ошибочных связей.", 
            font=ctk.CTkFont(size=11), 
            text_color="#8A8A8A",
            justify="left"
        )
        lbl_fuzzy_hint.grid(row=5, column=0, columnspan=2, padx=15, pady=(0, 15), sticky="w")
        
        # --- Кнопка Сохранить ---
        btn_save = ctk.CTkButton(
            scroll_frame, 
            text="💾 Сохранить настройки", 
            fg_color="#0D47A1", 
            hover_color="#1565C0",
            command=self.save_custom_settings,
            width=200
        )
        btn_save.grid(row=3, column=0, pady=20, padx=10)
        
        # Переключаем доступность полей согласно простому режиму
        self.toggle_simple_mode_ui()
        
    def toggle_simple_mode_ui(self):
        is_simple = self.var_simple_mode.get()
        state = "disabled" if is_simple else "normal"
        
        self.spn_hotel_margin.configure(state=state)
        self.chk_enable_id.configure(state=state)
        self.chk_enable_exact.configure(state=state)
        self.chk_enable_fuzzy.configure(state=state)
        self.spn_fuzzy_threshold.configure(state=state)
        
    def change_theme(self, new_theme: str):
        self.settings['theme'] = new_theme
        save_settings(self.settings)
        ctk.set_appearance_mode(new_theme)
        
        # Обновляем стили Treeview под новую тему
        self.style_treeviews()
        
        # Обновляем цвета текстового поля справки
        self.update_help_box_colors()
        
        # Обновляем сетки, если данные уже загружены
        if self.matches or self.unmatched_tp or self.unmatched_bt:
            self.populate_grids()
            
    def update_help_box_colors(self):
        if not hasattr(self, 'help_box') or not self.help_box:
            return
            
        is_dark = (self.settings.get('theme', 'Dark') == 'Dark')
        if is_dark:
            bg_color = "#1E1E1E"
            fg_color = "#CCCCCC"
            h1_color = "#90CAF9"
            h2_color = "#A5D6A7"
            bold_color = "#FFFFFF"
        else:
            bg_color = "#FFFFFF"
            fg_color = "#333333"
            h1_color = "#0D47A1"
            h2_color = "#2E7D32"
            bold_color = "#000000"
            
        self.help_box.configure(bg=bg_color, fg=fg_color)
        self.help_box.tag_config("h1", foreground=h1_color)
        self.help_box.tag_config("h2", foreground=h2_color)
        self.help_box.tag_config("bold", foreground=bold_color)
        self.help_box.tag_config("body", foreground=fg_color)
        self.help_box.tag_config("bullet", foreground=fg_color)
        
    def save_custom_settings(self):
        try:
            margin = float(self.var_hotel_margin.get())
            threshold = float(self.var_fuzzy_threshold.get())
            
            if not (0 <= margin <= 100):
                messagebox.showerror("Ошибка", "Процент маржи отелей должен быть от 0 до 100.")
                return
            if not (0 <= threshold <= 100):
                messagebox.showerror("Ошибка", "Порог нечеткого поиска должен быть от 0 до 100.")
                return
                
            self.settings['hotel_margin'] = margin
            self.settings['fuzzy_threshold'] = threshold
            self.settings['enable_id_match'] = bool(self.var_enable_id.get())
            self.settings['enable_exact_match'] = bool(self.var_enable_exact.get())
            self.settings['enable_fuzzy_match'] = bool(self.var_enable_fuzzy.get())
            self.settings['simple_mode'] = bool(self.var_simple_mode.get())
            self.settings['theme'] = self.var_theme.get()
            
            save_settings(self.settings)
            from models import ServiceItem
            ServiceItem.hotel_margin = margin
            
            messagebox.showinfo("Успешно", "Настройки сверки сохранены и будут применены при следующем расчете!")
        except ValueError:
            messagebox.showerror("Ошибка", "Пожалуйста, введите корректные числовые значения для маржи и порога.")
        
    def browse_tp(self):
        file = filedialog.askopenfilename(
            initialdir=self.default_folder,
            title="Выберите реестр TicketProf (или сводный файл)",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if file:
            self.input_file_tp.set(file)
            self.default_folder = os.path.dirname(file)
            self.settings['default_folder'] = self.default_folder
            save_settings(self.settings)
            
            fname = os.path.basename(file)
            self.lbl_tp_check.configure(text=f"✅ {fname[:20]}...")
            
            base = os.path.splitext(file)[0]
            self.output_excel.set(f"{base}_сопоставлено.xlsx")
            self.output_word.set(f"{base}_акт_сверки.docx")
            self.check_ready_state()
            
    def browse_bt(self):
        file = filedialog.askopenfilename(
            initialdir=self.default_folder,
            title="Выберите реестр Bars Tour (приходы)",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if file:
            self.input_file_bt.set(file)
            self.default_folder = os.path.dirname(file)
            self.settings['default_folder'] = self.default_folder
            save_settings(self.settings)
            
            fname = os.path.basename(file)
            self.lbl_bt_check.configure(text=f"✅ {fname[:20]}...")
            self.check_ready_state()
            
    def check_ready_state(self):
        if self.input_file_tp.get():
            self.btn_run.configure(state="normal")
            
    def start_analysis(self):
        tp_path = self.input_file_tp.get()
        
        if not tp_path or not os.path.exists(tp_path):
            messagebox.showerror("Ошибка", "Пожалуйста, выберите существующий файл сверки.")
            return
            
        self.btn_run.configure(state="disabled")
        self.lbl_status.configure(text="Чтение и парсинг файлов Excel...")
        self.progress_bar.set(0.2)
        
        threading.Thread(target=self.run_reconciliation_thread, args=(tp_path, None), daemon=True).start()
        
    def run_reconciliation_thread(self, tp_path: str, bt_path: str = None):
        try:
            from models import ServiceItem
            ServiceItem.hotel_margin = self.settings.get('hotel_margin', 10.0)
            
            tp_items, bt_items = load_data(tp_path, None)
            self.tp_items = tp_items
            self.bt_items = bt_items
            
            self.lbl_status.configure(text="Запущен интеллектуальный поиск...")
            self.progress_bar.set(0.5)
            
            self.matches, self.unmatched_tp, self.unmatched_bt = match_records(
                tp_items, bt_items, self.manual_links, self.settings
            )
            
            self.summary = calculate_reconciliation(self.tp_items, self.bt_items, self.matches)
            
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
        
        self.kpi_rate.update_value(f"{self.summary.total_tp_sum:,.2f} руб.")
        self.kpi_profit.update_value(f"{self.summary.total_bt_sum:,.2f} руб.")
        self.kpi_discrepancy.update_value(f"{self.summary.total_profit:,.2f} руб.")
        
        # Скрываем empty state и отображаем Treeview
        if self.empty_state_frame:
            self.empty_state_frame.grid_remove()
        self.tree_all.grid()
        
        self.populate_grids()
        
    def populate_grids(self):
        for tree in [self.tree_all, self.tree_mismatches, self.tree_tp, self.tree_bt, self.tree_manual_tp, self.tree_manual_bt]:
            for row in tree.get_children():
                tree.delete(row)
                
        is_dark = (self.settings.get('theme', 'Dark') == 'Dark')
        
        # Цвета выделения строк в зависимости от темы
        if is_dark:
            matched_bg, matched_fg = "#1E3E20", "#A5D6A7"
            unmatched_bg, unmatched_fg = "#3E2723", "#FFAB91"
            discrepancy_bg, discrepancy_fg = "#4A148C", "#E1BEE7"
        else:
            matched_bg, matched_fg = "#E2EFDA", "#225522"
            unmatched_bg, unmatched_fg = "#FCE4D6", "#883311"
            discrepancy_bg, discrepancy_fg = "#FFC7CE", "#9C0006"
            
        self.tree_all.tag_configure("matched", background=matched_bg, foreground=matched_fg)
        self.tree_all.tag_configure("unmatched", background=unmatched_bg, foreground=unmatched_fg)
        self.tree_all.tag_configure("discrepancy", background=discrepancy_bg, foreground=discrepancy_fg)
        
        self.tree_mismatches.tag_configure("matched", background=matched_bg, foreground=matched_fg)
        self.tree_mismatches.tag_configure("unmatched", background=unmatched_bg, foreground=unmatched_fg)
        self.tree_mismatches.tag_configure("discrepancy", background=discrepancy_bg, foreground=discrepancy_fg)
        
        # 1. Сначала заполняем сопоставления
        for tp, bt, method, score in self.matches:
            status_text = tp.get_status_text(bt)
            tag = "discrepancy" if status_text in ["Нетипичная маржа", "Несовпадение по суммам"] else "matched"
            
            tp_ids = list(tp.ids) if tp.ids else []
            bt_ids = list(bt.ids) if bt.ids else []
            all_ids = sorted(list(set(tp_ids + bt_ids)))
            tp_id = ", ".join(all_ids) if all_ids else "N/A"
            
            vals = (
                tp_id, tp.service_type,
                tp.doc, tp.desc, f"{tp.allocated_amount:,.2f}",
                bt.doc, bt.desc, f"{bt.amount:,.2f}",
                f"{(bt.amount - tp.allocated_amount):,.2f}",
                method, status_text
            )
            self.tree_all.insert("", "end", values=vals, tags=(tag,))
            
            if tag == "discrepancy":
                self.tree_mismatches.insert("", "end", values=vals, tags=(tag,))
            
        # 2. Затем нераспределенные из TicketProf
        for tp in self.unmatched_tp:
            tp_ids = sorted(list(tp.ids)) if tp.ids else []
            tp_id = ", ".join(tp_ids) if tp_ids else "N/A"
            status_text = tp.get_status_text(None)
            
            vals = (
                tp_id, tp.service_type,
                tp.doc, tp.desc, f"{tp.allocated_amount:,.2f}",
                "", "Отсутствует в Bars Tour", "0.00", "0.00",
                "Не сопоставлено", status_text
            )
            self.tree_all.insert("", "end", values=vals, tags=("unmatched",))
            self.tree_mismatches.insert("", "end", values=vals, tags=("unmatched",))
            
        # 3. Наконец, нераспределенные из Bars Tour
        for bt in self.unmatched_bt:
            bt_ids = sorted(list(bt.ids)) if bt.ids else []
            bt_id = ", ".join(bt_ids) if bt_ids else "N/A"
            status_text = bt.get_status_text(None)
            
            vals = (
                bt_id, bt.service_type,
                "", "Отсутствует в TicketProf", "0.00",
                bt.doc, bt.desc, f"{bt.amount:,.2f}", "0.00",
                "Не сопоставлено", status_text
            )
            
            if status_text == "Норма (Сбор в БТ)":
                self.tree_all.insert("", "end", values=vals, tags=("matched",))
            else:
                self.tree_all.insert("", "end", values=vals, tags=("unmatched",))
                self.tree_mismatches.insert("", "end", values=vals, tags=("unmatched",))
            
        # 4. Вкладка "В Тикете, нет в Барсе"
        for tp in self.unmatched_tp:
            tp_ids = sorted(list(tp.ids)) if tp.ids else []
            tp_id = ", ".join(tp_ids) if tp_ids else "N/A"
            self.tree_tp.insert("", "end", values=(
                tp.row, tp.date, tp.doc, tp.desc, tp.service_type, f"{tp.allocated_amount:,.2f}", tp_id
            ))
            self.tree_manual_tp.insert("", "end", values=(tp.row, tp.desc, f"{tp.allocated_amount:,.2f}"))
            
        # 5. Вкладка "В Барсе, нет в Тикете"
        for bt in self.unmatched_bt:
            bt_ids = sorted(list(bt.ids)) if bt.ids else []
            bt_id = ", ".join(bt_ids) if bt_ids else "N/A"
            self.tree_bt.insert("", "end", values=(
                bt.row, bt.date, bt.doc, bt.desc, bt.service_type, f"{bt.amount:,.2f}", bt_id
            ))
            self.tree_manual_bt.insert("", "end", values=(bt.row, bt.desc, f"{bt.amount:,.2f}"))
            
        self.update_manual_links_list()
        
    def update_manual_links_list(self):
        self.lst_links.configure(state="normal")
        self.lst_links.delete("1.0", "end")
        
        if not self.manual_links:
            self.lst_links.insert("end", "Нет ручных связей.\n")
        else:
            for tp_clean, bt_clean in self.manual_links.items():
                self.lst_links.insert("end", f"TP: «{tp_clean[:25]}...»\n ➔ BT: «{bt_clean[:25]}...»\n\n")
                
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
        
        details = []
        if len(values) == 11:
            details.append(f"Код идентификации: {values[0]}")
            details.append(f"Категория услуги: {values[1]}")
            details.append(f"--- TicketProf ---")
            details.append(f"Документ: {values[2]}")
            details.append(f"Описание: {values[3]}")
            details.append(f"Стоимость услуг: {values[4]} руб.")
            details.append(f"--- Bars Tour ---")
            details.append(f"Документ: {values[5]}")
            details.append(f"Описание: {values[6]}")
            details.append(f"Итого в Барсе: {values[7]} руб.")
            details.append(f"------------------")
            details.append(f"Прибыль: {values[8]} руб.")
            details.append(f"Метод привязки: {values[9]}")
            details.append(f"Статус: {values[10]}")
        else:
            details.append(f"Строка в Excel: {values[0]}")
            details.append(f"Дата транзакции: {values[1]}")
            details.append(f"Документ: {values[2]}")
            details.append(f"Номенклатура: {values[3]}")
            details.append(f"Тип услуги: {values[4]}")
            details.append(f"Сумма: {values[5]} руб.")
            details.append(f"ID: {values[6]}")
            
        self.details_panel.show_details("\n".join(details))
        
    def create_manual_link(self):
        sel_tp = self.tree_manual_tp.selection()
        sel_bt = self.tree_manual_bt.selection()
        
        if not sel_tp or not sel_bt:
            messagebox.showwarning("Внимание", "Выберите по одной записи в левой и правой таблицах.")
            return
            
        tp_vals = self.tree_manual_tp.item(sel_tp[0])["values"]
        bt_vals = self.tree_manual_bt.item(sel_bt[0])["values"]
        
        tp_row = int(tp_vals[0])
        bt_row = int(bt_vals[0])
        
        tp_item = next(item for item in self.tp_items if item.row == tp_row)
        bt_item = next(item for item in self.bt_items if item.row == bt_row)
        
        self.manual_links[tp_item.clean_desc] = bt_item.clean_desc
        self.save_manual_links()
        
        self.start_analysis()
        messagebox.showinfo("Готово", "Связь успешно создана!")
        
    def delete_selected_link(self):
        if not self.manual_links:
            return
            
        confirm = messagebox.askyesno("Очистить связи", "Вы действительно хотите удалить ВСЕ ручные связи?")
        if confirm:
            self.manual_links.clear()
            self.save_manual_links()
            self.start_analysis()
            messagebox.showinfo("Готово", "Все ручные связи удалены.")
            
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
            messagebox.showinfo("Готово", f"Акт сверки успешно сохранен:\n{out_path}")
        except Exception as e:
            messagebox.showerror("Ошибка сохранения", f"Не удалось создать Акт сверки:\n{e}")

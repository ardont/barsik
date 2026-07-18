# -*- coding: utf-8 -*-
"""
Пользовательские виджеты интерфейса
"""

import customtkinter as ctk
from tkinter import ttk

class KPICard(ctk.CTkFrame):
    """
    Карточка KPI для отображения финансовых метрик
    """
    def __init__(self, parent, title: str, value: str, color: tuple = ("#EAEAEA", "#2B2B2B"), text_color: tuple = ("#000000", "#FFFFFF")):
        super().__init__(parent, fg_color=color, corner_radius=10)
        
        # Заголовок карточки
        self.title_label = ctk.CTkLabel(
            self, 
            text=title, 
            font=ctk.CTkFont(family="Arial", size=11, weight="bold"), 
            text_color=("#666666", "#8A8A8A")
        )
        self.title_label.pack(anchor="w", padx=15, pady=(12, 2))
        
        # Значение карточки
        self.value_label = ctk.CTkLabel(
            self, 
            text=value, 
            font=ctk.CTkFont(family="Arial", size=18, weight="bold"),
            text_color=text_color
        )
        self.value_label.pack(anchor="w", padx=15, pady=(0, 12))

    def update_value(self, new_value: str, color: str = None):
        self.value_label.configure(text=new_value)
        if color:
            self.value_label.configure(text_color=color)

class DetailsPanel(ctk.CTkFrame):
    """
    Интерактивная панель деталей выбранной операции
    """
    def __init__(self, parent):
        super().__init__(parent, corner_radius=10, fg_color=("#EAEAEA", "#2B2B2B"))
        
        # Заголовок панели
        title_lbl = ctk.CTkLabel(
            self, 
            text="🔍 Детализированная информация по выбранной позиции", 
            font=ctk.CTkFont(family="Arial", size=12, weight="bold"),
            text_color=("#000000", "#FFFFFF")
        )
        title_lbl.pack(anchor="w", padx=15, pady=(10, 5))
        
        # Основной текст
        self.info_text = ctk.CTkTextbox(
            self, 
            font=ctk.CTkFont(family="Arial", size=10),
            fg_color=("#FFFFFF", "#1E1E1E"),
            text_color=("#333333", "#CCCCCC"),
            height=80
        )
        self.info_text.pack(fill="both", expand=True, padx=15, pady=(0, 10))
        self.info_text.configure(state="disabled")

    def show_details(self, details_str: str):
        self.info_text.configure(state="normal")
        self.info_text.delete("1.0", "end")
        self.info_text.insert("1.0", details_str)
        self.info_text.configure(state="disabled")
        
    def clear(self):
        self.info_text.configure(state="normal")
        self.info_text.delete("1.0", "end")
        self.info_text.configure(state="disabled")

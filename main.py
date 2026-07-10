# -*- coding: utf-8 -*-
"""
Главная точка входа приложения «Умная сверка 3.0»
"""

import sys
import os

# Добавляем пути проекта в sys.path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gui.app import ReconciliationApp

def main():
    try:
        app = ReconciliationApp()
        app.mainloop()
    except Exception as e:
        import traceback
        error_msg = f"Критическая ошибка при запуске приложения:\n{e}\n\n{traceback.format_exc()}"
        print(error_msg)
        
        # Если Tkinter доступен, выводим красивую ошибку
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Ошибка запуска", error_msg)
        except Exception:
            pass
            
        input("\nНажмите Enter для выхода...")

if __name__ == "__main__":
    main()

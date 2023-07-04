import tkinter as tk
from tkinter import filedialog
import pandas as pd

# Функция для обработки файла Excel и перевода описаний
def process_excel_file():
    # Запрос пути к файлу Excel
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    
    if file_path:
        try:
            # Загрузка файла Excel
            df = pd.read_excel(file_path)
            
            # Создание словаря с переводами цветов
            color_dict = {
                'Dark blue': 'синий',
                # Добавьте остальные переводы цветов в соответствии с вашим файлом
            }
            
            # Функция для замены значений в столбце с описанием
            def translate_colors(description):
                for color in color_dict:
                    if color in description:
                        description = description.replace(color, color_dict[color])
                return description
            
            # Применение функции перевода к столбцу с описаниями
            df['Описание'] = df['Описание'].apply(translate_colors)
            
            # Сохранение измененного файла Excel на русском языке
            output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
            if output_path:
                df.to_excel(output_path, index=False)
                tk.messagebox.showinfo("Успех", "Преобразование завершено.")
            else:
                tk.messagebox.showinfo("Ошибка", "Не указан путь для сохранения файла.")
            
        except Exception as e:
            tk.messagebox.showerror("Ошибка", str(e))
    else:
        tk.messagebox.showinfo("Ошибка", "Файл не выбран.")

# Создание окна приложения
window = tk.Tk()
window.title("Преобразование файла Excel")
window.geometry("300x150")

# Создание кнопки для выбора файла
file_button = tk.Button(window, text="Выбрать файл", command=process_excel_file)
file_button.pack(pady=20)

# Запуск главного цикла приложения
window.mainloop()
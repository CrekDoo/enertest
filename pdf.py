import os
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import shutil
import json
from tkinter import Tk, Label, Button, Frame, filedialog, messagebox, Entry, Checkbutton, IntVar

# Файл для сохранения пути к Excel
config_file = "config.json"

# Функция для загрузки пути к Excel из файла конфигурации
def load_excel_path():
    if os.path.exists(config_file):
        with open(config_file, "r") as f:
            config = json.load(f)
            return config.get("excel_path", "")
    return ""

# Функция для сохранения пути к Excel в файл конфигурации
def save_excel_path(path):
    with open(config_file, "w") as f:
        json.dump({"excel_path": path}, f)

def split_pdf(file_path, pages_per_split, output_dir):
    reader = PdfReader(file_path)
    total_pages = len(reader.pages)
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    for start in range(0, total_pages, pages_per_split):
        end = min(start + pages_per_split, total_pages)
        writer = PdfWriter()

        for page in range(start, end):
            writer.add_page(reader.pages[page])

        output_file = os.path.join(output_dir, f"part_{start // pages_per_split + 1}.pdf")
        with open(output_file, "wb") as f:
            writer.write(f)
        print(f"Сохранено: {output_file}")

def rename_and_move_pdfs(excel_file, output_dir):
    df = pd.read_excel(excel_file, sheet_name="pdf")
    
    for index, row in df.iterrows():
        new_name = row['B']
        destination_path = row['C']

        original_file = os.path.join(output_dir, f"part_{index + 1}.pdf")
        if os.path.exists(original_file):
            if not os.path.exists(destination_path):
                # Запрос на создание папки, если она не существует
                create_folder = messagebox.askyesno("Создание папки", f"Папка '{destination_path}' не существует. Хотите создать её?")
                if create_folder:
                    os.makedirs(destination_path)
                    print(f"Создана папка: {destination_path}")
                else:
                    print(f"Пропущено перемещение файла: {original_file}")
                    continue  # Пропускаем перемещение, если пользователь отказался

            new_file_path = os.path.join(destination_path, f"{new_name}.pdf")
            shutil.copy2(original_file, new_file_path)
            print(f"Скопировано и перемещено: {original_file} -> {new_file_path}")
            os.remove(original_file)
            print(f"Удален оригинальный файл: {original_file}")
        else:
            print(f"Файл не найден: {original_file}")

def select_pdf_file():
    file_path = filedialog.askopenfilename(title="Выберите PDF файл", filetypes=[("PDF files", "*.pdf")])
    if file_path:
        pdf_file_label.config(text=file_path)
    else:
        messagebox.showwarning("Предупреждение", "PDF файл не выбран.")

def select_excel_file():
    file_path = filedialog.askopenfilename(title="Выберите Excel файл", filetypes=[("Excel files", "*.xlsm")])
    if file_path:
        excel_file_label.config(text=file_path)
        if remember_var.get():
            save_excel_path(file_path)  # Сохраняем путь, если галочка установлена
    else:
        messagebox.showwarning("Предупреждение", "Excel файл не выбран.")

def process_files():
    pdf_file_path = pdf_file_label.cget("text")
    excel_file_path = excel_file_label.cget("text")

    if not pdf_file_path or not os.path.exists(pdf_file_path):
        messagebox.showwarning("Ошибка", "Пожалуйста, выберите корректный PDF файл.")
        return

    if not excel_file_path or not os.path.exists(excel_file_path):
        messagebox.showwarning("Ошибка", "Пожалуйста, выберите корректный Excel файл.")
        return

    try:
        pages_per_split = int(pages_per_split_entry.get())
        if pages_per_split <= 0:
            raise ValueError("Количество страниц должно быть положительным числом.")
    except ValueError as e:
        messagebox.showwarning("Ошибка", str(e))
        return

    output_directory = "output_pdfs"

    split_pdf(pdf_file_path, pages_per_split, output_directory)
    rename_and_move_pdfs(excel_file_path, output_directory)

    # Удаление папки output_pdfs после завершения обработки
    if os.path.exists(output_directory):
        shutil.rmtree(output_directory)
        print(f"Удалена папка: {output_directory}")

    messagebox.showinfo("Успех", "Обработка завершена!")

# Создаем главное окно
root = Tk()
root.title("PDF Splitter and Renamer")

# Создаем фрейм для элементов управления
frame = Frame(root)
frame.pack(padx=150, pady=10)

# Метки для отображения выбранных файлов
pdf_file_label = Label(frame, text="PDF файл не выбран", wraplength=300)
pdf_file_label.pack(pady=5)

excel_file_label = Label(frame, text="Excel файл не выбран", wraplength=300)
excel_file_label.pack(pady=5)

# Поле ввода для количества страниц на часть
Label(frame, text="Количество страниц на копию:").pack(pady=5)
pages_per_split_entry = Entry(frame)
pages_per_split_entry.insert(0, "2")  # Значение по умолчанию
pages_per_split_entry.pack(pady=5)

# Чекбокс для запоминания пути к Excel файлу
remember_var = IntVar()
remember_checkbox = Checkbutton(frame, text="Запомнить путь к Excel файлу", variable=remember_var)
remember_checkbox.pack(pady=5)

# Кнопки для выбора файлов
select_pdf_button = Button(frame, text="Выбрать PDF файл", command=select_pdf_file)
select_pdf_button.pack(pady=5)

select_excel_button = Button(frame, text="Выбрать Excel файл", command=select_excel_file)
select_excel_button.pack(pady=5)

# Загрузка сохраненного пути к Excel файлу
saved_excel_path = load_excel_path()
if saved_excel_path:
    excel_file_label.config(text=saved_excel_path)
    remember_var.set(1)  # Устанавливаем галочку, если путь загружен

# Кнопка для обработки файлов
process_button = Button(frame, text="Обработать файлы", command=process_files)
process_button.pack(pady=10)

# Запускаем главный цикл
root.mainloop()

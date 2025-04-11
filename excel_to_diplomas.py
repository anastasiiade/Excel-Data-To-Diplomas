import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from pathlib import Path
import fitz  # PyMuPDF
from pathlib import Path
import shutil

def add_text_to_pdf(input_pdf_path, output_pdf_path, text_data):
    """
    Добавляет текст в PDF-файл.

    Args:
        input_pdf_path (str): Путь к исходному PDF-файлу.
        output_pdf_path (str): Путь для сохранения измененного PDF-файла.
        text_data (dict): Словарь с текстом для добавления (ключи: координаты, значения: текст).
    """
    try:
        # Открываем PDF-файл
        doc = fitz.open(input_pdf_path)

        # Получаем первую страницу
        page = doc[0]
        font = fitz.Font("helv")  # this is a full font file
        page.insert_font(fontname="F0", fontbuffer=font.buffer)
        # Добавляем текст на страницу
        for (x, y), text in text_data.items():
            page.insert_text((x, y), text,
                             fontsize=19,
                             fontname='F0', color=(0.8, 0, 0))

        # Сохраняем изменения
        doc.save(output_pdf_path)
        doc.close()

        print(f"Результат сохранён в {output_pdf_path}")

    except Exception as e:
        print(f"Произошла ошибка: {e}")

def read_excel_rows(file_path):
    """
    Читает данные из Excel-файла и возвращает список строк.

    Аргументы:
    file_path (str): Путь к Excel-файлу.

    Возвращает:
    list: Список строк из Excel-файла.
    """
    try:
        # Читаем Excel-файл в DataFrame
        df = pd.read_excel(file_path)

        # Преобразуем DataFrame в список строк
        rows = df.values.tolist()

        return rows
    except FileNotFoundError:
        return "Файл не найден"
    except Exception as e:
        return f"Произошла ошибка: {str(e)}"

file_path = 'list.xlsx'  # Путь к файлу
rows = read_excel_rows(file_path)

if isinstance(rows, list):
    for row in rows:
        titul = str(row[0])
        stepen = str(int(float(row[1])))
        partisip = str(row[2])
        fio = str(row[3])
        klass = str(row[4])
        date = str(row[5])
        teacher = str(row[6])
        profi = str(row[7])
        input_pdf_path = 'frame.pdf'  # Путь к PDF-шаблону
        output_pdf_path = f"{fio}.pdf"  # Путь для сохранения изменённого PDF-файла

        # Текст для добавления в PDF (координаты (x, y) и текст)
        text_data = {
            (85, 260): f"{titul}" + f" {stepen} степени",
            (85, 282): f"{partisip}",
            (85, 380): f"{fio}",
            (85, 410): f" {klass} класс",
            (85, 620): f"{date}",
            (85, 490): f"{teacher}",
            (315, 500): f"{profi}",
        }

        # Добавляем текст в PDF
        add_text_to_pdf(input_pdf_path, output_pdf_path, text_data)
else:
    print(rows)  # Выводим сообщение об ошибке, если она произошла
import shutil

import openpyxl
import os

import pandas as pd
import json
import requests
from django.shortcuts import render, redirect, get_object_or_404
from django.http import FileResponse, Http404
from .forms import UploadFileForm
from .models import UploadedFile
from django.conf import settings


def upload_and_list_files(request):
    """Объединённая view: загрузка + список загруженных/обработанных."""

    if request.method == "POST":
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            form.save()
            return redirect("home")
    else:
        form = UploadFileForm()

    # Отфильтровываем реально существующие
    existing_files = [
        f for f in UploadedFile.objects.all()
        if os.path.exists(f.file.path)
    ]

    # Обработанные — ищем по имени "_final"
    uploads_dir = os.path.join(settings.BASE_DIR, "uploads")
    processed_files = []
    if os.path.exists(uploads_dir):
        for filename in os.listdir(uploads_dir):
            if "_final" in filename and filename.endswith(".xlsx"):
                processed_files.append(filename)

    return render(request, "processor/index.html", {
        "form": form,
        "files": existing_files,
        "processed_files": processed_files
    })


def upload_file(request):
    if request.method == "POST":
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            form.save()
            return redirect("file_list")  # Перенаправляем на список загруженных файлов
    else:
        form = UploadFileForm()
    return render(request, "processor/upload.html", {"form": form})


def process_xlsx(file_path):
    """Открывает Excel-файл, удаляет скрытые строки и пустые строки, затем сохраняет новый файл."""
    wb = openpyxl.load_workbook(file_path)

    for sheet in wb.worksheets:
        # Удаляем скрытые строки
        hidden_rows = [row for row in sheet.row_dimensions if sheet.row_dimensions[row].hidden]
        for row in sorted(hidden_rows, reverse=True):
            sheet.delete_rows(row)

    # Сохраняем временный файл
    temp_file_path = file_path.replace(".xlsx", "_temp.xlsx")
    wb.save(temp_file_path)

    # Читаем и фильтруем через pandas (удаляем строки, где все значения NaN)
    df = pd.read_excel(temp_file_path)
    df = df.dropna(how="all")  # Удаляем строки, где все значения пустые
    df = df.dropna(axis=1, how="all") # Удаляем столбцы, где ВСЕ значения пустые

    # Удаляем временный файл
    os.remove(temp_file_path)

    # Сохраняем финальный результат
    new_file_path = file_path.replace(".xlsx", "_processed.xlsx")
    df.to_excel(new_file_path, index=False)

    return new_file_path


def file_list(request):
    """Выводит только актуальные загруженные файлы + обработанные."""

    # Фильтрация: показываем только те записи, у которых файл существует физически
    all_uploaded = UploadedFile.objects.all()
    existing_uploaded_files = [
        f for f in all_uploaded if os.path.exists(f.file.path)
    ]

    # Обработанные файлы из папки /uploads/
    uploads_dir = os.path.join(settings.BASE_DIR, "uploads")
    processed_files = []
    if os.path.exists(uploads_dir):
        for filename in os.listdir(uploads_dir):
            if "_final" in filename and filename.endswith(".xlsx"):
                processed_files.append(filename)

    return render(request, "processor/file_list.html", {
        "files": existing_uploaded_files,
        "processed_files": processed_files,
    })


def convert_xlsx_to_text(file_path):
    """Читает .xlsx, конвертирует в JSON (текстовый формат)."""
    df = pd.read_excel(file_path)
    return df.to_json(orient="records", force_ascii=False, indent=4)  # JSON с сохранением символов


def convert_text_to_xlsx(json_text, output_path):
    """Принимает JSON и сохраняет его в .xlsx."""
    try:
        data = json.loads(json_text)
        df = pd.DataFrame(data)
        df.to_excel(output_path, index=False)
        return True
    except Exception as e:
        print(f"Ошибка преобразования: {e}")
        return None


def apply_priorities_from_chatgpt(original_path, chatgpt_path):
    """
    Обновляет колонку 'Приоритет' в оригинальном файле, используя данные из ChatGPT,
    где есть колонки '№' и 'Приоритет'.
    """
    original_df = pd.read_excel(original_path)
    chatgpt_df = pd.read_excel(chatgpt_path)

    # Проверим, что нужные колонки есть
    if "№" not in original_df.columns or "Приоритет" not in chatgpt_df.columns:
        print("Ошибка: отсутствует нужная колонка в одном из файлов.")
        return None

    # Создаём маппинг: № → Приоритет
    priority_mapping = chatgpt_df.set_index("№")["Приоритет"].to_dict()

    # Обновляем значения в оригинальной таблице
    original_df["Приоритет"] = original_df["№"].map(priority_mapping).combine_first(original_df["Приоритет"])

    # Сохраняем новый файл
    final_path = original_path.replace(".xlsx", "_final.xlsx")
    original_df.to_excel(final_path, index=False)
    return final_path


def send_to_chatgpt(text_data, prompt):
    """Отправляет текст в ChatGPT API и получает обработанный JSON."""
    url = "https://api.proxyapi.ru/openai/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {os.getenv('PROXYAPI_KEY')}",
        "Content-Type": "application/json",
    }

    messages = [
        {"role": "developer", "content": f"{prompt}"},
        {"role": "user", "content": f"Данные:\n{text_data}"}
    ]

    payload = {
        "model": "gpt-4o-mini",
        "messages": messages,
    }

    response = requests.post(url, headers=headers, json=payload)

    if response.status_code == 200:
        return response.json()["choices"][0]["message"]["content"]

    return None


def process_with_chatgpt(request, file_id):
    """Обрабатывает файл, сохраняет результат и обновляет страницу."""
    uploaded_file = get_object_or_404(UploadedFile, id=file_id)
    original_file_path = uploaded_file.file.path
    processed_file_path = process_xlsx(original_file_path)

    text_data = convert_xlsx_to_text(processed_file_path)
    prompt = settings.CHATGPT_PROMPT
    processed_text = send_to_chatgpt(text_data, prompt)

    if processed_text:
        chatgpt_path = original_file_path.replace(".xlsx", "_chatgpt.xlsx")
        result = convert_text_to_xlsx(processed_text, chatgpt_path)

        if result:
            final_file_path = apply_priorities_from_chatgpt(original_file_path, chatgpt_path)

            if final_file_path:
                # Копируем файл в uploads/ с _final
                uploads_dir = os.path.join(settings.BASE_DIR, "uploads")
                os.makedirs(uploads_dir, exist_ok=True)

                final_filename = os.path.basename(final_file_path).replace("_processed", "_final")
                final_path = os.path.join(uploads_dir, final_filename)
                shutil.copy(final_file_path, final_path)

    # Возвращаемся на главную страницу
    return redirect("home")

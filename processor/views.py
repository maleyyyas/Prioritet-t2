import openpyxl
import os

import pandas as pd
import json
import requests
from django.shortcuts import render, redirect, get_object_or_404
from django.http import FileResponse
from .forms import UploadFileForm
from .models import UploadedFile
from django.conf import settings


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
    """Выводит список загруженных файлов с кнопками обработки."""
    files = UploadedFile.objects.all()
    return render(request, "processor/file_list.html", {"files": files})


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
        return output_path
    except Exception as e:
        print(f"Ошибка преобразования: {e}")
        return None


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
    """Отправляет обработанный файл в ChatGPT и позволяет скачать результат."""
    uploaded_file = get_object_or_404(UploadedFile, id=file_id)
    processed_file_path = process_xlsx(uploaded_file.file.path)

    text_data = convert_xlsx_to_text(processed_file_path)

    prompt = settings.CHATGPT_PROMPT
    processed_text = send_to_chatgpt(text_data, prompt)

    if processed_text:
        new_file_path = uploaded_file.file.path.replace(".xlsx", "_chatgpt.xlsx")
        saved_file_path = convert_text_to_xlsx(processed_text, new_file_path)

        if saved_file_path:
            return FileResponse(open(saved_file_path, "rb"), as_attachment=True, filename=os.path.basename(saved_file_path))

    return render(request, "processor/error.html", {"message": "Ошибка обработки AI."})

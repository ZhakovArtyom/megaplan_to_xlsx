import html  # для декодирования HTML-кодов
import logging
import os
import tempfile
import time
from datetime import datetime
from typing import List, Dict

import requests
from fastapi import APIRouter, HTTPException
from fastapi.responses import JSONResponse
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

from config import settings

router = APIRouter()

MEGAPLAN_API_URL = settings.MEGAPLAN_API_URL
MEGAPLAN_API_KEY = settings.MEGAPLAN_API_KEY
MEGAPLAN_HEADER = {
    "Authorization": f"Bearer {MEGAPLAN_API_KEY}",
    "Content-Type": "application/json"
}

# Словарь с русскими названиями месяцев
MONTHS_RU = {
    1: 'января',
    2: 'февраля',
    3: 'марта',
    4: 'апреля',
    5: 'мая',
    6: 'июня',
    7: 'июля',
    8: 'августа',
    9: 'сентября',
    10: 'октября',
    11: 'ноября',
    12: 'декабря'
}


def get_project_issues(project_id: str, url: str, header: Dict[str, str]) -> List[Dict]:
    url = f"{url}/api/v3/project/{project_id}/issues"
    try:
        response = requests.get(url, headers=header, timeout=120)
        response.raise_for_status()
        project_data = response.json()["data"]
        logging.info(f"Получены задачи проекта с ID: {project_id}")
        time.sleep(1)
        return project_data
    except requests.exceptions.RequestException as e:
        logging.exception(f"Error occurred while getting project {project_id}: {e}")
        raise


def get_project(project_id: str, url: str, header: Dict[str, str]) -> Dict:
    url = f"{url}/api/v3/project/{project_id}"
    try:
        response = requests.get(url, headers=header, timeout=120)
        response.raise_for_status()
        project_data = response.json()["data"]
        logging.info(f"Получен проект с ID: {project_id}")
        time.sleep(1)
        return project_data
    except requests.exceptions.RequestException as e:
        logging.exception(f"Error occurred while getting project {project_id}: {e}")
        raise


def get_task(task_id: str, url: str, header: Dict[str, str]) -> Dict:
    url = f"{url}/api/v3/task/{task_id}"
    try:
        response = requests.get(url, headers=header, timeout=120)
        response.raise_for_status()
        task_data = response.json()["data"]
        logging.info(f"Получена задача с ID: {task_id}")
        time.sleep(1)
        return task_data
    except requests.exceptions.RequestException as e:
        logging.exception(f"Error occurred while getting task {task_id}: {e}")
        raise


def get_comment(comment_id: str, url: str, header: Dict[str, str]) -> str:
    url = f"{url}/api/v3/comment/{comment_id}"
    try:
        response = requests.get(url, headers=header, timeout=120)
        response.raise_for_status()
        comment_data = response.json()["data"]
        logging.info(f"Получен комментарий с ID: {comment_id}")
        time.sleep(1)
        return comment_data["content"]
    except requests.exceptions.RequestException as e:
        logging.exception(f"Error occurred while getting comment {comment_id}: {e}")
        return ""


def is_product(text: str) -> bool:
    return any(char.isdigit() for char in text)


def clean_html(text: str) -> str:
    """Удаление HTML-тегов и декодирование символов."""
    text = html.unescape(text)  # Декодируем HTML-символы
    text = text.replace('<br />', '\n').replace('</p>', '\n').replace('<p>', '').replace('</strong>', '').replace(
        '<strong>', '')
    return text.strip()


def process_tasks(project_name: str, issues: List[Dict], sheet) -> None:
    row = 2
    for issue in issues:
        issue_name = issue["name"]
        issue_data = get_task(issue["id"], MEGAPLAN_API_URL, MEGAPLAN_HEADER)

        development_task = next(
            (task for task in issue_data["subTasks"] if task["name"] == "1 💡 Разработка продуктов"), None)
        if development_task:
            logging.info(f'Получена задача 1 💡 Разработка продуктов с ID {development_task["id"]}')
            development_task_data = get_task(development_task["id"], MEGAPLAN_API_URL, MEGAPLAN_HEADER)

            # Декодирование и очистка данных
            products_raw = development_task_data["subject"]
            products_clean = clean_html(products_raw)

            products = products_clean.split("\n\n")  # Разделение продуктов

            for product in products:
                if is_product(product):
                    sheet.cell(row=row, column=1, value=row - 1).alignment = Alignment(horizontal="center",
                                                                                       vertical="center")
                    sheet.cell(row=row, column=2, value=project_name).alignment = Alignment(horizontal="center",
                                                                                            vertical="center")
                    sheet.cell(row=row, column=3, value=issue_name).alignment = Alignment(horizontal="center",
                                                                                          vertical="center")

                    sheet.cell(row=row, column=4, value=product).alignment = Alignment(wrap_text=True,
                                                                                       horizontal="left",
                                                                                       vertical="center")

                    # Форматируем дату
                    raw_date = development_task_data["actualStart"]["value"]
                    date_obj = datetime.strptime(raw_date, "%Y-%m-%dT%H:%M:%S%z")
                    formatted_date = f"{date_obj.day} {MONTHS_RU[date_obj.month]}"
                    sheet.cell(row=row, column=5, value=formatted_date).alignment = Alignment(horizontal="center",
                                                                                              vertical="center")

                    sheet.cell(row=row, column=6,
                               value=development_task_data["responsible"]["name"]).alignment = Alignment(
                        horizontal="center", vertical="center")
                    sheet.cell(row=row, column=7, value=issue_data["owner"]["name"]).alignment = Alignment(
                        horizontal="center", vertical="center")

                    raw_materials_task = next(
                        (task for task in development_task_data["subTasks"] if task["name"] == "1. Поставщики сырья"),
                        None)
                    if raw_materials_task:
                        logging.info(f'Получена задача 1. Поставщики сырья с ID {raw_materials_task["id"]}')
                        raw_materials_task_data = get_task(raw_materials_task["id"], MEGAPLAN_API_URL, MEGAPLAN_HEADER)
                        if raw_materials_task_data["comments"]:
                            raw_materials_comment = get_comment(raw_materials_task_data["comments"][0]["id"],
                                                                MEGAPLAN_API_URL,
                                                                MEGAPLAN_HEADER)
                            sheet.cell(row=row, column=8, value=raw_materials_comment).alignment = Alignment(
                                horizontal="center", vertical="center")

                    packaging_task = next(
                        (task for task in development_task_data["subTasks"] if
                         task["name"] == "2. Поставщики упаковки"),
                        None)
                    if packaging_task:
                        logging.info(f'Получена задача 2. Поставщики упаковки с ID {packaging_task["id"]}')
                        packaging_task_data = get_task(packaging_task["id"], MEGAPLAN_API_URL, MEGAPLAN_HEADER)
                        if packaging_task_data["comments"]:
                            packaging_comment = get_comment(packaging_task_data["comments"][0]["id"], MEGAPLAN_API_URL,
                                                            MEGAPLAN_HEADER)
                            sheet.cell(row=row, column=9, value=packaging_comment).alignment = Alignment(
                                horizontal="center", vertical="center")

                    if development_task_data["comments"]:
                        last_comment = get_comment(development_task_data["comments"][-1]["id"], MEGAPLAN_API_URL,
                                                   MEGAPLAN_HEADER)
                        sheet.cell(row=row, column=10, value=last_comment).alignment = Alignment(horizontal="center",
                                                                                                 vertical="center")

                    row += 1


def upload_file(filename, real_name):
    url = f"{MEGAPLAN_API_URL}/api/file"
    headers = {
        "Authorization": f"Bearer {MEGAPLAN_API_KEY}"

    }
    real_name = real_name = f"{real_name}.{filename.rsplit('.', 1)[-1]}"
    with open(filename, 'rb') as file:
        files = {'files[]': (real_name, file, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
        try:
            response = requests.post(url, headers=headers, files=files)
            response.raise_for_status()
            file_data = response.json()['data'][0]
            file_id = file_data['id']
            return file_id
        except requests.RequestException as e:
            logging.exception(f"Error uploading file: {e}")
            return None
        finally:
            os.remove(filename)


@router.get("/test")
async def test_endpoint():
    return JSONResponse(status_code=200, content={"message": "Test request successful!"})


@router.post("/unloading-tasks/{project_id}")
async def unload_tasks(project_id: str):
    # Создаем временный файл Excel
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Задачи проекта"

        # Настройка стилей для первой строки
        header_fill = PatternFill(start_color="FFEB84", end_color="FFEB84", fill_type="solid")  # Жёлтый цвет
        header_font = Font(bold=True)
        header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        # Границы для разграничения заголовков
        thin_border = Border(left=Side(style='thin'),
                             right=Side(style='thin'),
                             top=Side(style='thin'),
                             bottom=Side(style='thin'))

        # Установка заголовков и применение стиля ко всей первой строке
        headers = ["№", "Бренд", "Линейка", "Наименование", "Дата запуска работы", "Ответственный БМ",
                   "Ответственный ОЗ", "Сырье", "Упаковка", "Примечание"]
        for col_num, header in enumerate(headers, 1):
            cell = sheet.cell(row=1, column=col_num, value=header)
            cell.fill = header_fill  # Применяем жёлтый цвет
            cell.font = header_font
            cell.alignment = header_alignment
            cell.border = thin_border  # Добавляем границы

        # Получение данных проекта
        project_issues = get_project_issues(project_id, MEGAPLAN_API_URL, MEGAPLAN_HEADER)
        project_data = get_project(project_id, MEGAPLAN_API_URL, MEGAPLAN_HEADER)
        project_name = project_data["name"]

        # Запуск обработки задач
        process_tasks(project_name, project_issues, sheet)

        # Настройка ширины колонок
        for column_cells in sheet.columns:
            length = max(len(str(cell.value)) for cell in column_cells) + 2  # Добавляем дополнительную ширину
            sheet.column_dimensions[column_cells[0].column_letter].width = length

        sheet.column_dimensions['A'].width = 5
        sheet.column_dimensions['D'].width = 50

        workbook.save(tmp.name)
        tmp.close()

    # Загружаем файл
    file_id = upload_file(tmp.name, real_name=project_name)

    if not file_id:
        raise HTTPException(status_code=500, detail="Error uploading file")

    # Отправляем комментарий с файлом
    url = f"{MEGAPLAN_API_URL}/api/v3/project/{project_id}/comments"
    headers = {
        'Authorization': f'Bearer {MEGAPLAN_API_KEY}',
        'Content-Type': 'application/json'
    }

    body = {
        "contentType": "CommentCreateActionRequest",
        "comment": {
            "contentType": "Comment",
            "content": f"Задачи проекта {project_name}",
            "attaches": [
                {
                    "id": file_id,
                    "contentType": "File"
                }
            ],
            "subject": {
                "id": project_id,
                "contentType": "Project"
            }
        },
        "transports": [
            {}
        ]
    }

    response = requests.post(url, headers=headers, json=body)
    response.raise_for_status()

    return JSONResponse(status_code=200, content={"message": "Tasks unloaded successfully"})

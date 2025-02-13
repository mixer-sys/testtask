import time
import os

from dotenv import load_dotenv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import smtplib
import xlwings as xw


BOOK_NAME = 'TestTask1.xlsx'
LIST_NAME = 'Sheet1'
FILE_PATH = 'C:\\Dev\\testtask\\TestTask2.xlsx'
DRIVER_PATH = 'C:\\Windows\\System32\\yandexdriver.exe'
BROWSER_PATH = ('C:\\Users\\user\\AppData\\Local\\Yandex\\YandexBrowser\\'
                'Application\\browser.exe')
LETTER_THEME = 'Список тем для доклада'
RECEPIENT = 'den.s1m@yandex.ru'
Theme_Sources = []

load_dotenv()

username = os.getenv('USERNAME')
password = os.getenv('PASSWORD')


def send_mail(username=username, password=password,
              recepient=RECEPIENT, file_path=FILE_PATH):

    smtp_server = 'smtp.yandex.com'
    smtp_port = 587
    username = username
    password = password

    msg = MIMEMultipart()
    msg['From'] = username
    msg['To'] = recepient
    msg['Subject'] = LETTER_THEME

    body = LETTER_THEME
    msg.attach(MIMEText(body, 'plain'))

    filename = file_path
    attachment = open(filename, 'rb')

    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header(
        'Content-Disposition',
        f'attachment; filename={filename.split("/")[-1]}'
        )
    msg.attach(part)

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(username, password)
        server.send_message(msg)
        print("Письмо отправлено успешно!")
    except Exception as e:
        print(f"Ошибка: {e}")
    finally:
        server.quit()
        attachment.close()


def task1(book_name=BOOK_NAME, list_name=LIST_NAME):
    """
    В ОС Windows открыт Excel файл "TestTask1.xlsx".
    В файле на листе "Sheet1" есть три столбца - "Task", "Status" и "Comment"
    (наполнение таблицы случайное, в столбце "Status" два варианта заполнения
    "Done" и "In progress").
    С помощью библиотеки xlwings получить доступ к данному файлу и раскрасить
    строки в зеленый цвет если статус "Done"
    и в красный цвет если статус - "In progress".
    Сохранить файл в конце работы.
    """
    app = xw.apps.active

    workbook = app.books[book_name]

    sheet = workbook.sheets[list_name]

    statuses = sheet.range('B1:B100').value
    for i, status in enumerate(statuses):
        if status == 'Done':
            # RGB для зеленого цвета
            sheet.range(f'A{i + 1}:Z{i + 1}').color = (0, 255, 0)
        if status == 'In progress':
            # RGB для красного цвета
            sheet.range(f'A{i + 1}:Z{i + 1}').color = (255, 0, 0)

    workbook.save()


def task2(file_path=FILE_PATH, list_name=LIST_NAME):
    """
    1. В ОС Windows есть Excel файл "TestTask2.xlsx" в папке
    "С:\\Documents\\Reports"
    (или другой папке на выбор). В файле на листе "Sheet1" есть
    два столбца - "Theme", и "Sources"
    (в столбце "Theme" случайное наполнение, столбец "Sources" - пуст).
    Собрать все темы, указанные в файле в список для поиска.
    2. С помощью автоматизации открыть Яндекс браузер (или Chrome браузер),
    перейти на сайт "ya.ru (http://ya.ru/)"
    и сделать поисковый запрос для каждой темы. Взять со страницы
    с результатами поиска первые 3 ссылки из раздела "Результаты поиска",
    продублировать название темы в файле и вставить каждую найденную ссылку
    в столбец "Sources" на отдельную строку.
    В строке с заголовками таблицы включить возможность фильтрации данных.
    Сохранить файл в конце работы.
    3. Полученный файл выслать в качестве приложения к письму с темой
    "Список тем для доклада"
    на указанный адрес электронной почты. Для отправки использовать
    сервис яндекс почты.
    """

    wb = xw.Book(file_path)
    sheet = wb.sheets[list_name]

    themes = sheet.range('A:A').value
    themes = [
        theme for theme in themes if theme is not None and theme != 'Theme'
    ]

    options = Options()
    options.binary_location = BROWSER_PATH
    service = Service(DRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=options)

    for theme in themes:

        driver.get("https://ya.ru")
        time.sleep(15)
        search_box = driver.find_element(By.NAME, "text")
        search_box.send_keys(theme)

        search_button = driver.find_element(By.XPATH,
                                            "//button[@type='submit']")
        search_button.click()

        time.sleep(15)

        search_results = driver.find_element(By.ID, "search-result")

        links = search_results.find_elements(By.TAG_NAME, "a")

        for i, link in enumerate(links[:3], start=1):
            Theme_Sources.append([theme, link.get_attribute('href')])

        sheet.range("A2").value = Theme_Sources
        sheet.range("A1").expand().api.AutoFilter(1)
        wb.save()

    wb.close()


if __name__ == "__main__":
    task1()
    task2()
    send_mail()

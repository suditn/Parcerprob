from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.edge.options import Options
import re
import pandas as pd
from bs4 import BeautifulSoup
import shutil
import os
from pathlib import Path
import requests
import logging
from concurrent.futures import ThreadPoolExecutor

# Логи
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Поиск слова
pattern = re.compile(r'\w+')

# Настройка браузера ссылок и путей файлов
options = Options()
options.headless = True
driver = webdriver.Edge(options=options)
url = 'https://www.vishay.com/en/inductors/'
save_path = str(Path(__file__).parent.resolve())
img_small_save_path = os.path.join(save_path, "image", "small_inductors")
datasheet_save_path = os.path.join(save_path, "Datasheet")
headers = {'User-Agent': "scrapping_script/1.0"}

# Создание папок
def create_directories():
    Path(img_small_save_path).mkdir(parents=True, exist_ok=True)
    Path(datasheet_save_path).mkdir(parents=True, exist_ok=True)

# Выгрузка полной таблицы на страницу
def get_web(u):
    driver.get(u)
    logging.info(f'Открыта страница: {u}')
    create_directories()
    wait = WebDriverWait(driver, 10)
    option2 = driver.find_element('xpath', '/html/body/div[1]/div/div[2]/div[2]/div[2]/div/div/div/div/div[3]/div[2]/label/select/option[1]')
    max_entries = driver.find_element('xpath', '/html/body/div[1]/div/div[2]/div[2]/div[2]/div/div/div/div/div[3]/div[1]/div').text
    driver.execute_script('arguments[0].value = arguments[1]', option2, pattern.findall(max_entries)[5])
    option2.click()
    return driver.page_source

# Функция для скачивания файлов.
def download_file(url, path, headers=None):
    try:
        response = requests.get(url, headers=headers, stream=True)
        file_dir = os.path.dirname(path)
        Path(file_dir).mkdir(parents=True, exist_ok=True)
        if not os.path.exists(path):
            with open(path, 'wb') as out_file:
                shutil.copyfileobj(response.raw, out_file)
            logging.info(f'Файл {os.path.basename(path)} успешно загружен и сохранен.')
        else:
            logging.info(f'Файл {os.path.basename(path)} уже существует.')
    except Exception as e:
        logging.error(f'Ошибка при загрузке файла {url}: {e}')
    del response

# Функция для обработки HTML и запуска параллельной загрузки.
def process_html(html_source):
    soup = BeautifulSoup(html_source, "lxml")
    table = soup.find('table', {'id': 'poc'})
    images = table.findAll('img')
    columns = [i.get_text(strip=True) for i in table.find_all("th")]
    data = [[td.get_text(strip=True) for td in tr.find_all("td")] for tr in table.find("tbody").find_all("tr")]

    df = pd.DataFrame(data, columns=columns)
    img_src = []
    datasheet_src = []
    download_tasks = []
    previous_img_src = ''
    previous_datasheet_src = ''
    i = 0


    # ThreadPoolExecutor для параллельной загрузки.
    with ThreadPoolExecutor(max_workers=5) as executor:
    # Разгрузка
        for img in images:
            series = df['Series▲▼'][i]
            if img['src'].split('/')[-2] == 'pt-small':
                img_filename = img['alt'] + '.png'
                img_path = os.path.join(img_small_save_path, img_filename)
                img_src.append(img_path)
                if previous_img_src != img['src'] and img['alt'] != "Datasheet":
                    download_tasks.append(executor.submit(download_file, 'https://www.vishay.com/' + img['src'], img_path))
                    previous_img_src = img['src']

                datasheet_filename = series + '.pdf'
                datasheet_path = os.path.join(datasheet_save_path, series, datasheet_filename)
                datasheet_src.append(datasheet_path)
                i += 1
                if previous_datasheet_src != series and img['alt'] != "Datasheet":
                    download_tasks.append(executor.submit(download_file, 'https://www.vishay.com/doc?' + img['alt'], datasheet_path, headers))
                    previous_datasheet_src = series

    # Ожидаем завершения всех задач загрузки.
    for task in download_tasks:
        task.result()

    return df, img_src, datasheet_src

# Функция для сохранения данных в Excel.
def save_to_excel(df, img_src, datasheet_src, save_path, url):
    excel_path = os.path.join(save_path, url.split('/')[-2] + '.xlsx')
    with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
        df_img = pd.DataFrame(img_src, columns=['Product Image'])
        df_datasheet = pd.DataFrame(datasheet_src, columns=['Datasheet'])
        df_final = df.join(df_img, lsuffix='_left', rsuffix='_right').join(df_datasheet, rsuffix='_datasheet')
        df_final.to_excel(writer, index=False, sheet_name='Inductors')
        worksheet = writer.sheets['Inductors']
        worksheet.autofit()

# Проверка работает ли всё
try:
    web_source = get_web(url)
    df, img_src, datasheet_src = process_html(web_source)
    save_to_excel(df, img_src, datasheet_src, save_path, url)
    logging.info('Данные успешно сохранены.')
finally:
    driver.quit()
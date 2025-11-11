from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
from datetime import datetime
from collections import defaultdict
from dotenv import load_dotenv
import pandas
import os
import argparse


def get_year_form(year):
    if year % 10 == 1 and year % 100 != 11:
        return "год"
    elif 2 <= year % 10 <= 4 and (year % 100 < 10 or year % 100 >= 20):
        return "года"
    else:
        return "лет"


def get_excel_data(file_path):
    excel_data = pandas.read_excel(
        file_path,
        na_values='',
        keep_default_na=False
    )
    excel_data = excel_data.fillna('')
    wines = excel_data.to_dict('records')
    return wines


def get_wines_category(wines):
    wines_by_category = defaultdict(list)
    for wine in wines:
        category = wine['Категория']
        wines_by_category[category].append(wine)
    wines_by_category = dict(wines_by_category)
    return wines_by_category


def main():
    load_dotenv()
    parser = argparse.ArgumentParser(description='Генератор сайта вин')
    parser.add_argument(
        '--excel-file',
        type=str,
        default=os.getenv('WINE_EXCEL_FILE', 'assortment_of_wine.xlsx'),
        help='Путь к Excel файлу с данными о винах'
    )
    args = parser.parse_args()
    if not os.path.exists(args.excel_file):
        print(f"Ошибка: Файл '{args.excel_file}' не найден!")
        print("Доступные способы указания пути:")
        print("1. Аргумент командной строки: --excel-file путь/к/файлу.xlsx")
        print("2. Переменная окружения: WINE_EXCEL_FILE=путь/к/файлу.xlsx")
        print("3. Файл .env: WINE_EXCEL_FILE=путь/к/файлу.xlsx")
        print("4. Значение по умолчанию: assortment_of_wines.xlsx")
        return

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    wines_data = get_excel_data(args.excel_file)
    template = env.get_template('template.html')
    established_year = 1920
    years_count = datetime.now().year - established_year
    year_form = get_year_form(years_count)
    rendered_page = template.render(
        years=years_count,
        year_form=year_form,
        wines=get_wines_category(wines=wines_data)
    )
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)
    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()

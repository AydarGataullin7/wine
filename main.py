from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
from datetime import datetime
from collections import defaultdict
import pandas


def get_year_form(year):
    if year % 10 == 1 and year % 100 != 11:
        return "год"
    elif 2 <= year % 10 <= 4 and (year % 100 < 10 or year % 100 >= 20):
        return "года"
    else:
        return "лет"


def get_excel_data():
    excel_data = pandas.read_excel(
        'wine3.xlsx', na_values='', keep_default_na=False)
    excel_data = excel_data.fillna('')
    wines = excel_data.to_dict('records')
    return wines


def wines_category(wines):
    wines_by_category = defaultdict(list)
    for wine in wines:
        category = wine['Категория']
        wines_by_category[category].append(wine)
    wines_by_category = dict(wines_by_category)
    return wines_by_category


def main():
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    wines_data = get_excel_data()
    template = env.get_template('template.html')
    years_count = datetime.now().year - 1920
    year_form = get_year_form(years_count)
    rendered_page = template.render(
        years=years_count,
        year_form=year_form,
        wines=wines_category(wines=wines_data)
    )
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)
    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()

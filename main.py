from http.server import HTTPServer, SimpleHTTPRequestHandler
from unicodedata import name
from jinja2 import Environment, FileSystemLoader, select_autoescape
import datetime
import openpyxl
from collections import defaultdict


def get_company_age(company_founding_year):
    now_year = datetime.datetime.today().year
    return now_year - company_founding_year


def get_age_word_ending(age):
    word_ending_rules = {
        "лет": [11, 12, 13, 14],
        "год": [1],
        "года": [2, 3, 4]
    }
    for word_ending, digits in word_ending_rules.items():
        if (age % 100) in digits:
            return word_ending
        elif (age % 10) in digits:
            return word_ending
    return "лет"


def parse_raw_catalog(excel_catalog_path):
    excel_obj = openpyxl.load_workbook(excel_catalog_path)
    sheet = excel_obj.active
    category_column = 0
    name_column = 1
    sort_column = 2
    price_column = 3
    img_path_column = 4
    stock_column = 5
    indent_row = 2
    wine_catalog = defaultdict(list)
    for row in sheet.iter_rows(indent_row):
        category = row[category_column].value
        name = row[name_column].value
        sort = row[sort_column].value
        price = row[price_column].value
        img_path = row[img_path_column].value
        stock = row[stock_column].value
        wine_catalog[category].append({
            "name": name,
            "sort": sort,
            "price": price,
            "img_path": "images/{}".format(img_path),
            "stock": stock,
        })
    return wine_catalog


if __name__ == "__main__":
    env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.get_template('template.html')
    raw_product_catalog = 'wine3.xlsx'
    product_catalog = parse_raw_catalog(raw_product_catalog)
    company_founding_year = 1920
    company_age = get_company_age(company_founding_year)
    age_word_ending = get_age_word_ending(company_age)
    rendered_page = template.render(
        wine_catalog = product_catalog,
        company_age = company_age,
        age_word_ending = age_word_ending
    )
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)
    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()

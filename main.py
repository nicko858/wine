import datetime
from collections import defaultdict
from http.server import HTTPServer, SimpleHTTPRequestHandler
from unicodedata import category

import openpyxl
from jinja2 import Environment, FileSystemLoader, select_autoescape


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


def get_grouped_product_catalog(excel_catalog_path):
    excel_obj = openpyxl.load_workbook(excel_catalog_path)
    sheet = excel_obj.active
    columns_names_idx = 0
    product_data_slice = slice(1,None)
    all_rows = [row for row in sheet.iter_rows()]
    product_rows = all_rows[product_data_slice]
    columns_names = [row.value for row in all_rows[columns_names_idx]]
    grouped_product_catalog = defaultdict(list)
    for row in product_rows:
        column_values = [cell.value for cell in row]
        category = column_values[0]
        grouped_product_catalog[category].append(
            dict(
                zip(
                    columns_names[product_data_slice],
                    column_values[product_data_slice]
                    )
                )
        )
    return grouped_product_catalog


if __name__ == "__main__":
    env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.get_template('template.html')
    product_catalog_path = 'wine3.xlsx'
    product_catalog = get_grouped_product_catalog(product_catalog_path)
    company_founding_year = 1920
    company_age = get_company_age(company_founding_year)
    age_word_ending = get_age_word_ending(company_age)
    rendered_page = template.render(
        product_catalog = product_catalog,
        company_age = company_age,
        age_word_ending = age_word_ending
    )
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)
    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()

import datetime
from collections import defaultdict
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas as pd
from jinja2 import Environment, FileSystemLoader, select_autoescape


def get_company_age(company_founding_year):
    now_year = datetime.datetime.today().year
    return now_year - company_founding_year


def get_age_word_ending(age):
    word_ending_rules = {
        'лет': [11, 12, 13, 14],
        'год': [1],
        'года': [2, 3, 4],
    }
    for word_ending, digits in word_ending_rules.items():
        if (age % 100) in digits:
            return word_ending
        elif (age % 10) in digits:
            return word_ending
    return 'лет'


def get_grouped_product_catalog(product_catalog_path):
    products = pd.read_excel(
        product_catalog_path,
        na_values=None,
        keep_default_na=False,
    ).to_dict(orient='records')
    grouped_products = defaultdict(list)
    for product in products:
        grouped_products[product['Категория']].append(product)
    return grouped_products


if __name__ == '__main__':
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml']),
    )
    template = env.get_template('template.html')
    product_catalog_path = 'product_catalog.xlsx'
    product_catalog = get_grouped_product_catalog(product_catalog_path)
    company_founding_year = 1920
    company_age = get_company_age(company_founding_year)
    age_word_ending = get_age_word_ending(company_age)
    rendered_page = template.render(
        product_catalog=product_catalog,
        company_age=company_age,
        age_word_ending=age_word_ending,
    )
    with open('index.html', 'w', encoding='utf8') as file_handler:
        file_handler.write(rendered_page)
    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()

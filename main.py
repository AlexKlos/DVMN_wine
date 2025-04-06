from collections import defaultdict
from datetime import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

from jinja2 import Environment, FileSystemLoader, select_autoescape
import pandas as pd

try:
    import settings
    try:
        WINE_DATA_PATH = settings.WINE_DATA_PATH
    except (AttributeError):
        WINE_DATA_PATH = 'wine.xlsx'
    try:
        TEMPLATE_PATH = settings.TEMPLATE_PATH
    except (AttributeError):
        TEMPLATE_PATH = 'template.html'
except (ImportError):
    WINE_DATA_PATH = 'wine.xlsx'
    TEMPLATE_PATH = 'template.html'


def pluralize_years(years):
    if 11 <= years % 100 <= 14:
        return 'лет'
    last_digit = years % 10

    if last_digit == 1:
        return 'год'
    elif 2 <= last_digit <= 4:
        return 'года'
    
    return 'лет'


def main():
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    
    template = env.get_template(TEMPLATE_PATH)
  
    winery_age = datetime.now().year - 1920
    
    wine_data = pd.read_excel(WINE_DATA_PATH, 
                              sheet_name='Лист1', 
                              na_values='', 
                              keep_default_na=False,
                              dtype=str).fillna('').to_dict(orient='records')
    
    wine_dict = defaultdict(list)
    
    for wine in wine_data:
        category = wine.get('Категория', 'Без категории')
        wine_dict[category].append(wine)
    
    rendered_page = template.render(
        winery_age = winery_age,
        years = pluralize_years(winery_age),
        wine_data = wine_dict,
    )
    
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)
        
    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
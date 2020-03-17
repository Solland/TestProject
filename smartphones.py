import requests
import openpyxl
from bs4 import BeautifulSoup

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'
                  'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.132 Safari/537.36'
}
url = requests.get('https://technopoint.ru/catalog/recipe/e351231ca6161134/2020-goda/', headers=headers)
soup = BeautifulSoup(url.text, "html.parser")


def main():
    codes = get_product_codes()
    titles, links = get_titles()
    prices, images = get_prices_and_images(links)
    save_as_xls(titles, codes, prices, images)


def get_product_codes():
    '''Возвращает список, состоящий
    из кодов продукта
    '''
    product_codes = soup.find_all('div', {'class': 'product-info__code'})
    product_codes = [item.text for item in product_codes][0:10]
    return [int(s) for s in product_codes]


def get_titles():
    '''Возвращает списки названий и ссылок
    на каждую конкретную модель
    '''
    product_titles = soup.find_all('a', {'class': 'ui-link', 'data-role': 'clamped-link'})[0:10]
    model_names = [item.text for item in product_titles]
    links = [f'https://technopoint.ru{item.get("href")}' for item in product_titles]
    return model_names, links


def get_prices_and_images(links):
    '''Возвращает списки цен и ссылок
    на изображения в формате jpg
    '''
    prices = []
    images = []

    for i in range(len(links)):
        html = requests.get(f'{links[i]}', headers=headers).text
        model = BeautifulSoup(html, "html.parser")
        prices.append(model.find(attrs={'class': "current-price-value"}).get('data-price-value'))
        images.append(model.find(attrs={'class': "lightbox-img"}).get('href'))
        prices = [int(s) for s in prices]
    return prices, images


def save_as_xls(model_names, product_codes, prices, images):
    '''Сохраняет файл, как документ excel
    '''
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.column_dimensions['A'].width = 55
    sheet.column_dimensions['B'].width = 12
    sheet.column_dimensions['D'].width = 150
    sheet['A1'] = 'Наименование модели'
    sheet['B1'] = 'Код товара'
    sheet['C1'] = 'Цена'
    sheet['D1'] = 'Изображение'
    for i in range(2, 12):
        sheet.append([model_names[i - 2], product_codes[i - 2], prices[i - 2], images[i - 2]])
    wb.save('smartphones.xls')


if __name__ == "__main__":
    main()

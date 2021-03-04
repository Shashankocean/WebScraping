from bs4 import BeautifulSoup
from openpyxl import Workbook
import requests


def get_product_information(url):
    html_ = requests.session().get(url)
    return html_


if __name__ == '__main__':
    # Collecting list of mens air shoes
    nike_web_new_arr = requests.session().get("https://www.nike.com/in/w/mens-air-force-1-shoes-5sj3yznik1zy7ok")
    test = BeautifulSoup(nike_web_new_arr.content, 'html.parser')
    new_arrival = test.find_all(class_='product-card')

    # Creating table in excel
    excel_file = Workbook()
    sheet_new_arrival = excel_file.active
    sheet_new_arrival.title = 'New Arrival Shoes'
    sheet_new_arrival['A1'] = 'Name'
    sheet_new_arrival['B1'] = 'Shot Info'
    sheet_new_arrival['C1'] = 'Price'
    sheet_new_arrival['D1'] = 'Discount'
    sheet_new_arrival['E1'] = 'Colors'
    sheet_new_arrival['F1'] = 'Product Link'

    # collecting shoes information like name, price, discount price and short information
    for n in range(0, len(new_arrival)):
        product_link_ = new_arrival[n].find('figure').find('a').get('href')
        name_ = new_arrival[n].find('figure').find(class_='product-card__title').text
        short_info_ = new_arrival[n].find('figure').find(class_='product-card__subtitle').text
        price_ = new_arrival[n].find('figure').find(attrs={"data-test": "product-price"}).text
        discount_price = new_arrival[n].find('figure').find(attrs={"data-test": 'product-price-reduced'})
        discount_price_ = ''
        if discount_price:
            discount_price_ = discount_price.text
        sheet_new_arrival['A{}'.format(n + 2)] = name_
        sheet_new_arrival['B{}'.format(n + 2)] = short_info_
        sheet_new_arrival['C{}'.format(n + 2)] = price_
        sheet_new_arrival['D{}'.format(n + 2)] = discount_price_
        sheet_new_arrival['F{}'.format(n + 2)] = product_link_

        # Colour information is not present in listing page hence Opening each shoe's detail page for color information.
        html = get_product_information(product_link_)
        colours_ = []
        any_color = BeautifulSoup(html.content, 'html.parser').find(class_='colorway-images-wrapper')
        colour = any_color.find_all('a') if any_color else 0
        if colour:
            print(colour)
            for m in range(0, len(colour)):
                colours_.append(colour[m].find('img').get('alt'))
        else:
            single_color = BeautifulSoup(html.content, 'html.parser').find(
                class_='description-preview__color-description')
            print('One Color: {}'.format(single_color))
            colours_.append(single_color.text) if single_color else ''

        # finally storing colour information in table as well
        sheet_new_arrival['E{}'.format(n + 2)] = '|'.join(colours_)

    excel_file.save('Nike.xlsx')

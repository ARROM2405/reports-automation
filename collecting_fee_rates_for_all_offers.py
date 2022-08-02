import os
from pathlib import Path

import dotenv

import requests
import bs4

env_path = Path('.', '.env')
dotenv.load_dotenv(dotenv_path=env_path)

session = requests.Session()
session.post(os.getenv('platform_login_link'),
             data={'in_user': os.getenv('platform_username'), 'in_pass': os.getenv('platform_password')})
pagination_num = 1
table_row_text = None
while table_row_text != 'Нет элементов для отображения':
    total_list_response = session.get(f'{os.getenv("offers_link")}?page={pagination_num}')
    total_list_soup = bs4.BeautifulSoup(total_list_response.text, 'lxml')
    table = total_list_soup.find_all('table', {'class': 'table no-margin table-condensed table-bordered'})[0]
    tbody = table.find_all('tbody')[0]
    trs = tbody.find_all('tr')
    table_row_text = trs[0].text
    if table_row_text == 'Нет элементов для отображения':
        break
    for tr in trs:
        tds = tr.find_all('td')
        offer_id = tds[1].text
        offer_name = tds[2].text
        offer_details = session.get(f'{os.getenv("offer_prices_link")}/{offer_id}')
        offer_details_soup = bs4.BeautifulSoup(offer_details.text, 'lxml')
        web_list = offer_details_soup.find_all('ul', {'id': 'offerprices'})
        print(f'ID: {offer_id}; Product: {offer_name}; Link: {os.getenv("offer_prices_link")}/{offer_id}')
        webs_divided = web_list[0].find_all('li', {'class': 'offer-price hold-578'})
        if len(webs_divided) > 0:
            for web in webs_divided:
                web_name = web.find('h3', {'class': 'panel-title'}).text.strip()
                if web_name:

                    agency_tariff_value = web.find('span', {'class': 'margin text-danger'})
                    if agency_tariff_value:
                        agency_tariff_value = agency_tariff_value.text.strip()
                    else:
                        print('No webs')
                        continue
                    company_tariff_value = web.find('span', {'class': 'margin text-success'}).text.strip()
                    referal_tariff_value = web.find('span', {'class': 'margin text-warning'})
                    if referal_tariff_value:
                        print(f'{web_name}: {agency_tariff_value.split(":")[1].strip()}, '
                              f'{company_tariff_value.split(":")[1].strip()},'
                              f' {referal_tariff_value.text.split(":")[1].strip()}')
                    else:
                        print(f'{web_name}: {agency_tariff_value.split(":")[1].strip()}, '
                              f'{company_tariff_value.split(":")[1].strip()}')

        else:
            print('No webs')
        print()
    pagination_num += 1

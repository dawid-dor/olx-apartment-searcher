import requests
import datetime
from bs4 import BeautifulSoup
import sys
from openpyxl import load_workbook
from datetime import datetime
import time

#TO_DO:
# 1. check links matching (details are not synced with proper hyperlink)
# 2. check for duplicates (i think link mismatch will fix that)
# 3. otodom module
# 4. workflow improvement

#GLOBALS
PAGE_NUMBER = 1
URL = 'https://www.olx.pl/nieruchomosci/mieszkania/warszawa/?search%5Bfilter_float_price%3Afrom%5D=1600&search%5Bfilter_float_price%3Ato%5D=2100&search%5Bfilter_enum_rooms%5D%5B0%5D=two&search%5Bprivate_business%5D=private'
OLX_LINKS = []
OTODOM_LINKS = []
OFFERS_LIST = []

months = [
    'stycznia', 'lutego', 'marca', 'kwietnia', 'maja', 'czerwca', 'lipca',
    'sierpnia', 'wrzesnia', 'pazdziernika', 'listopada', 'grudnia'
]


class Offer:
    def __init__(self, title, price, link, localization, time, date):
        self.title = title
        self.price = price
        self.link = link
        self.localization = localization
        self.time = time
        self.date = date

    def return_json_object(self):
        offer_details = {
            "title": self.title,
            "price": self.price,
            "link": self.link,
            "localization": self.localization,
            "time": self.time,
            "date": self.date
        }
        return offer_details


def olx_offer_parser(link):
    page_details = requests.get(link)
    soup_details = BeautifulSoup(page_details.content, 'html.parser')

    title = soup_details.find('div',
                              class_='offer-titlebox').find('h1').text.strip()
    try:
        localization = soup_details.find('address').find('p').text
    except AttributeError:
        localization = 'Warszawa'

    date_location = soup_details.find('em').find('strong').text
    date = [x.strip() for x in date_location.split(',')]
    date_time = date[0].replace("o ", "")
    date_array = date[1].split(" ")
    date_date = [
        str(months.index(x) + 1) if x in months else x for x in date_array
    ]
    date_date[1] = '0' + date_date[1] if len(
        date_date[1]) == 1 else date_date[1]
    date_date_formated = datetime.strptime("/".join(date_date), '%d/%m/%Y')

    price = int("".join([
        str(s) for s in [
            int(s) for s in soup_details.find(
                'div', class_='pricelabel').text.split() if s.isdigit()
        ]
    ]))
    #FIX THIS
    if price == 20:
        price = 2000
    subprice_check = soup_details.find('span', string="Czynsz (dodatkowo)")
    if subprice_check is not None:
        try:
            subprice = int("".join([
                str(s) for s in [
                    int(s) for s in subprice_check.parent.parent.findNext(
                        'strong').contents[0].split() if s.isdigit()
                ]
            ]))
        except ValueError:
            subprice = 0
    else:
        subprice = 0
    total = price + subprice
    offer = Offer(title, total, link, localization, date_time,
                  date_date_formated.date())

    return offer.return_json_object()


# loop through the offers and extract their links
for i in range(PAGE_NUMBER, 10):
    url = URL + '&page={}'.format(PAGE_NUMBER)
    page = requests.get(url)
    soup = BeautifulSoup(page.content, 'html.parser')
    offers = soup.find_all('div', class_='offer-wrapper')
    for offer in offers:
        offer_link = offer.find('a')['href'].strip()
        domain_check = 'olx' in offer_link
        if domain_check:
            OLX_LINKS.append(offer_link)
        else:
            OTODOM_LINKS.append(offer_link)
    PAGE_NUMBER += 1

# remove duplicates
OLX_LINKS = list(set(OLX_LINKS))
OTODOM_LINKS = list(set(OTODOM_LINKS))

for link in OLX_LINKS:
    offer_details = olx_offer_parser(link)
    OFFERS_LIST.append(offer_details)

# Worksheet initialization

workbook = load_workbook(filename=sys.argv[1])
sheet = workbook.active

sheet['A1'] = 'Tytuł'
sheet['B1'] = 'Lokalizacja'
sheet['C1'] = 'Cena'
sheet['D1'] = 'Godzina dodania'
sheet['E1'] = 'Data dodania'
sheet['F1'] = 'Link do ogłoszenia'

current_row = 2
for offer in OFFERS_LIST:
    #title
    sheet['A{}'.format(current_row)] = offer['title']
    #localization
    sheet['B{}'.format(current_row)] = offer['localization']
    #price
    sheet['C{}'.format(current_row)] = offer['price']
    #time
    sheet['D{}'.format(current_row)] = offer['time']
    #date
    sheet['E{}'.format(current_row)] = offer['date']
    #link
    sheet['F{}'.format(current_row)].hyperlink = offer['link']
    current_row += 1
workbook.save(filename=sys.argv[1])

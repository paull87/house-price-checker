# zooplaPrices.py - This script runs daily to get the current price of our
# property at DA7 6AR

import datetime as dt, requests, bs4, openpyxl as o, sys

# Get current date and split into day, month and year.
def get_date():
    """Returns the current date as a string"""
    fmt = '{0}/{1}/{2}'
    i = dt.datetime.now()
    day = ('0' + str(i.day))[-2:]
    month = ('0' + str(i.month))[-2:]
    year = i.year
    return fmt.format(day, month, year)

def download_page(site):
    """Downloads and parses the website, returning the bs4 result."""
    res = requests.get(site)
    if res.ok:
        soup = bs4.BeautifulSoup(res.text, 'html.parser')
        return soup
    return None

def get_price(soup):
    """Checks for price and returns the result."""
    try:
        node = soup.select('.big')
        return node[1].getText()
    except:
        return None

def open_sheet(xfile, name):
    """Opens the given excel spreadsheet and returns the given sheet."""
    wb = o.load_workbook(xfile)
    return wb.get_sheet_by_name(name)

def open_sheet(xfile, sheet_name, date, price):
    """Opens the excel spreadsheet and adds a new row to the excel spreadsheet for the latest price."""
    wb = o.load_workbook(xfile)
    sheet = wb.get_sheet_by_name(sheet_name)
    # Get last row
    nextcell = str(sheet.max_row + 1)
    # Write the new price
    write_record(sheet, nextcell, date, price)
    # Save the change
    wb.save(xfile)

def write_record(sheet, nextcell, date, price):
    """Adds a new row to the excel spreadsheet for the latest price."""
    sheet['A' + nextcell].value = date
    sheet['B' + nextcell].value = price
    sheet['B' + nextcell].number_format = '"£"#,##0;[Red]\\\\-"£"#,##0'

if __name__ == '__main__':

    # Set the file location
    xfile = 'C:\\Paul Personal\\House\\HousePrice.xlsx'
    sheet = 'Sheet1'
    # The zoopla web page with the prices
    zoopla = 'http://www.zoopla.co.uk/property/10-spring-vale/bexleyheath/da7-6ar/6092805'
    # Current date
    date = get_date()

    # Download the zoopla page
    soup = download_page(zoopla)
    if not soup:
        print('Unable to parse web page.')
        sys.exit()
    # Check for the price
    price = get_price(soup)
    if price:
        open_sheet(xfile, sheet, date, price)
        print('Latest price added.')







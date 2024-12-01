import argparse
import gspread
import datetime
import os
from secrets import *

import smtplib
from email.message import EmailMessage

from pprint import pprint as pp
from urllib.request import urlopen
from bs4 import BeautifulSoup
from babel.numbers import parse_decimal, get_currency_symbol
from oauth2client.service_account import ServiceAccountCredentials

# from secrets.py:
# email_sender = '...'
# email_password = '...'
# gsuite_credentials = {...}
# email_recipient = '...'
# gsheet_id = '...'
# debug_mode = false


root_url = "https://www.festoolrecon.com"

def convert_currency(value):
  code = get_currency_symbol('USD');
  return parse_decimal(value.strip(code), locale='en_US')

def send_email(insert_rows, update_rows, unchanged_rows):
  if not email_password or not email_sender:
    print('No email sender or password found.')
    return

  subject = 'changes in product inventory'
  if len(insert_rows) == 1:
    subject = insert_rows[0][2]
  elif len(insert_rows) > 0:
    subject = f'{len(insert_rows)} new products'
  elif len(unchanged_rows) > 0:
    subject = f'{len(insert_rows)} out-of-stock products'

  # create the message object
  msg = EmailMessage()
  msg['Subject'] = f'[Festool Recon] - {subject}'
  msg['From'] = email_sender
  msg['To'] = email_recipient
  msg.set_content(
    f'New products:  {''.join([f"\n\t{t[2]} - ${t[4]} / ${t[3]}" for t in insert_rows])}\n\n'
    f'Sold out products:  {''.join([f"\n\t{t[2]}" for t in unchanged_rows])}\n\n'
    f'Still on sale:  {''.join([f"\n\t{t[2]}" for t in update_rows])}')

  if args.debug:
    print(f'[Debug] Email message: \n{msg}')
    return

  # connect to the Gmail SMTP server and send the message
  with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
    smtp.login(email_sender, email_password)
    smtp.send_message(msg)
    print(f'Email sent from {email_sender}')

def extract_collection(soup):
  collection_link = soup.select_one('a.collection-card')
  if collection_link:
    url = collection_link.get('href')
    if not url.startswith('http'):
      url = f'{root_url}{url}'
    print(f'Collection page found, navigating to <{url}>')

    with urlopen(url) as collection_page_response:
      collection_page = BeautifulSoup(collection_page_response.read(), features="html.parser")
      if args.debug:
        print(f"[Debug] Multiple products:")
      products = []

      for card_el in collection_page.select('.product-card__info'):
        product_title = card_el.select_one('.product-card__name').getText().strip()
        original_price = convert_currency(card_el.select_one(".product-card__regular-price").getText().strip())
        sale_price = convert_currency(''.join(card_el.select_one(".product-card__price").findAll(string=True, recursive=False)).strip())
        if args.debug:
          print(f"    {product_title} - {original_price} - {sale_price}")
        products.append((product_title, original_price, sale_price))
      return products

  print('No products found...')
  return []

def write_to_google_sheets(products):
  if not products:
    print('No products to write...')
    return

  # Authenticate with Google Sheets API using credentials file
  scope = ["https://spreadsheets.google.com/feeds",
          'https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/drive.file',
          'https://www.googleapis.com/auth/drive']

  client = gspread.service_account_from_dict(gsuite_credentials)

  # Open the Google Sheet by ID
  sheet = client.open_by_key(gsheet_id).sheet1

  all_rows = sheet.get_all_values()[1:]

  first_row = all_rows[0]

  collection_rows = []

  # Take all product rows in the collection based on the date last modified, taking all products that are part of the grouping.
  for row in all_rows:
    if row[1] == all_rows[0][1]:
      collection_rows.append(row)
    else:
      break

  # Matching product rows
  new_product_names = [t[0] for t in products]
  matched_stored_products = [row for row in collection_rows if row[2] in new_product_names]
  matched_stored_product_names = [t[2] for t in matched_stored_products]

  current_date = datetime.datetime.now()
  current_datetime_str = current_date.strftime("%Y-%m-%d %H:%M:%S")

  insert_rows = [[current_datetime_str, current_datetime_str, p[0], float(p[1]), float(p[2])] for p in products if p[0] not in matched_stored_product_names]
  update_rows = [[t[0], current_datetime_str, *t[2:5]] for t in matched_stored_products]
  unchanged_rows = [t[0:5] for t in collection_rows if t[2] not in matched_stored_product_names]

  print(f'Updated rows: {''.join([f"\n  {t}" for t in update_rows])}')
  print(f'Inserting rows: {''.join([f"\n  {t}" for t in insert_rows])}')
  print(f'Unchanged rows: {''.join([f"\n  {t}" for t in unchanged_rows])}')

  if args.debug:
    print('Debug mode: skipping write to Google Sheets and sending email...')
  else:
    print(f'Writing {len(insert_rows)} new rows, {len(update_rows)} updated rows, {len(unchanged_rows)} unchanged rows')
    if len(update_rows) > 0:
      sheet.update(range_name = f'A2:E{2 + len(update_rows)}', values = update_rows)
    if len(unchanged_rows) > 0:
      sheet.update(range_name = f'A2:E{2 + len(update_rows) + len(unchanged_rows)}', values = unchanged_rows)
    # If there are new products, insert into sheet and send email.
    if len(insert_rows) > 0:
      sheet.insert_rows(insert_rows, 2)

  if len(unchanged_rows) > 0 or len(insert_rows) > 0:
    print('Sending email for product changes')
    send_email(insert_rows, update_rows, unchanged_rows)
  else:
    print('No product changes, not sending email')



# Create the argument parser
arg_parser = argparse.ArgumentParser(description="Festool recon site scraper.")

# Add the optional --debug flag (default is False)
arg_parser.add_argument(
  '--debug',       # The flag itself (optional)
  action='store_true',  # If the flag is provided, store True
  help='Enable debug mode'  # Description of what the flag does
)

# Parse the command-line arguments
args = arg_parser.parse_args()

# Check if --debug flag was provided
if args.debug:
    print("Running in  debug mode...")

with urlopen(root_url) as response:
  soup = BeautifulSoup(response.read(), features="html.parser")
  product_title_el = soup.select_one("h1.product-single__title")

  products = []

  if product_title_el:
    product_title = product_title_el.getText().strip()
    original_price = convert_currency(soup.select_one("#ComparePrice-product-template").getText().strip())
    sale_price = convert_currency(soup.select_one("#ProductPrice-product-template").getText().strip())
    products = [(product_title, original_price, sale_price)]

    if args.debug:
      print(f"Single product: {product_title} - {original_price} - {sale_price}")
  else:
    products = extract_collection(soup)

  write_to_google_sheets(products)


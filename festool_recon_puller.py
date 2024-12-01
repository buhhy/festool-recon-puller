import gspread
import datetime
import os
import secrets

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
# gsheet_id = '...'

def convert_currency(value):
  code = get_currency_symbol('USD');
  return parse_decimal(value.strip(code), locale='en_US')

def send_email(product_title, original_price, sale_price):
  if not email_password or not email_sender:
    print('No email sender or password found.')
    return

  # create the message object
  msg = EmailMessage()
  msg['Subject'] = f'[Festool Recon] - {product_title}'
  msg['From'] = email_sender
  msg['To'] = 'terence.lei@live.ca'
  msg.set_content(f'New product on Festool Recon: {product_title} - ${sale_price} / ${original_price}.')

  # connect to the Gmail SMTP server and send the message
  with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
    smtp.login(email_sender, email_password)
    smtp.send_message(msg)
    print(f'Email sent from {email_sender}')

url = "https://www.festoolrecon.com/"
with urlopen(url) as response:
  soup = BeautifulSoup(response.read(), features="html.parser")
  product_title = soup.select_one("h1.product-single__title").getText().strip()
  original_price = convert_currency(soup.select_one("#ComparePrice-product-template").getText().strip())
  sale_price = convert_currency(soup.select_one("#ProductPrice-product-template").getText().strip())

  # Authenticate with Google Sheets API using credentials file
  scope = ["https://spreadsheets.google.com/feeds",
          'https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/drive.file',
          'https://www.googleapis.com/auth/drive']

  client = gspread.service_account_from_dict(gsuite_credentials)

  # Open the Google Sheet by ID
  sheet = client.open_by_key(gsheet_id).sheet1

  first_row = sheet.row_values(2)

  current_date = datetime.datetime.now()
  current_datetime_str = current_date.strftime("%Y-%m-%d %H:%M:%S")

  if first_row[2] == product_title:
    # Item hasn't changed, update the last updated timestamp.
    sheet.update_cell(2, 2, current_datetime_str);
    print(f"{current_datetime_str}: updated line: {product_title} - ${original_price}/${sale_price}")
  else:
    # Write a new row.
    sheet.insert_row([current_datetime_str, current_datetime_str, product_title, float(original_price), float(sale_price)], index=2)
    # Send an email to alert new product.
    send_email(product_title, original_price, sale_price)
    print(f"{current_datetime_str}: wrote line: {product_title} - ${original_price}/${sale_price}")


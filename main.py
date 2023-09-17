from time import sleep
import yfinance as yf
from datetime import datetime
from polygon import ReferenceClient
import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from dotenv import load_dotenv
import logging


def main():
    load_dotenv()
    CLIENT = ReferenceClient(os.getenv("POLYGON_IO_API_KEY"))
    SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
    SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
    SHEET_NAME = "csv_data"
        
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        headers = ['ticker', 'current_price', 'ex_date', 'pay_date', 'cash', 'div_yield', 'frequency', 'update_time']
        sheet.values().update(spreadsheetId=SPREADSHEET_ID,
                                    range=f'{SHEET_NAME}!A{1}', valueInputOption = "USER_ENTERED", 
                                    body={"values": [headers]}).execute()
        stocks = sheet.values().get(spreadsheetId=SPREADSHEET_ID,range=f'{SHEET_NAME}!A2:A').execute()["values"]
        stocks = [''.join(x) for x in stocks]

        for count, stock in enumerate(stocks):    
            divData = [CLIENT.get_stock_dividends(stock)]
            for ticker in divData:
                d_ticker = ticker["results"][-1]["ticker"]
                d_current_price = yf.Ticker(stock).info['currentPrice']
                d_ex_dividend_date = ticker["results"][-1]["ex_dividend_date"]
                d_pay_date = ticker["results"][-1]["pay_date"]
                d_cash_amount = ticker["results"][-1]["cash_amount"]
                d_div_frequency = ticker["results"][-1]["frequency"]
                div_yield = calculate_div_yield(ticker, d_div_frequency, d_current_price)
                update_time = str(datetime.now())
                sheet.values().update(spreadsheetId=SPREADSHEET_ID,
                                    range=f'{SHEET_NAME}!A{count+2}', valueInputOption = "USER_ENTERED", 
                                    body={"values": [[d_ticker, d_current_price, d_ex_dividend_date, d_pay_date, d_cash_amount, div_yield, d_div_frequency, update_time]]}).execute()
            if ((count+1)%5 == 0):
                logging.info("polygon.io 5 request per minute exceeded. Sleep for 60 seconds.")
                sleep(60)
    except HttpError as err:
        print(err)

def calculate_div_yield(div_data, d_div_frequency, d_current_price, dividend_type="FWD"):
    """
    TTM: The sum of actual ordinary cash dividend payments in the last 12 months divided by the last close price
    FWD: Last paid dividend multiplid by frequency divided current price
    """
    if dividend_type == "TTM":
        sum_dividends = 0
        for data in div_data["results"][-d_div_frequency:]:
            sum_dividends += data["cash_amount"]
        return round(sum_dividends/d_current_price, 4)
    elif dividend_type == "FWD":
        return round(div_data["results"][-1]["cash_amount"]*d_div_frequency/d_current_price, 4)
    else:
        raise("Unknown dividend type, select 'TTM' or 'FWD'")

if __name__ == '__main__':
    main()
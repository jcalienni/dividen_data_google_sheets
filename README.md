# dividend_data_google_sheets
With this code you can get dividend data, and this data is uploaded automatically to specified google sheet  
The stock tickers are obtained from column A starting from A2  
![image](https://github.com/jcalienni/dividen_data_google_sheets/assets/53088875/eef750e4-ee66-437e-8fce-6ea5385e9e13)


Market data is taken from yahoo finance and polygon.io libraries  
Need polygon.io account. Free trial allows 5 api requests per minute  
Need google cloud and activate googlesheet API. need to generate credentials json file from google cloud  
Rename json credentials file to "credentials.json" and add it to project root  

# Requirements:  
create .venv  
connect to .venv  
pip install -r requirements.txt  
Inside .env file you should add the following information:  
SPREADSHEET_ID = "you_spread_sheet_id"  
POLYGON_IO_API_KEY = "your_polygon_io_api_key"  

# Run:  
connect to .venv  
run "python main.py"  


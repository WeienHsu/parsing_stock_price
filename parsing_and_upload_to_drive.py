import os, sys
os.environ['REQUESTS_CA_BUNDLE'] = os.path.join(os.path.dirname(sys.argv[0]), 'cacert.pem')
import pandas as pd
import requests
import xlsxwriter
from bs4 import BeautifulSoup
from bs4.dammit import EncodingDetector
from datetime import datetime, date
# for upload excel to google drive
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload


def isfloat(num):
    try:
        float(num)
        return True
    except ValueError:
        return False

def get_soup(url):    
    headers = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:78.0) Gecko/20100101 Firefox/78.0",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"}
    resp = requests.get(url, headers=headers)
    http_encoding = resp.encoding if 'charset' in resp.headers.get('content-type', '').lower() else None
    html_encoding = EncodingDetector.find_declared_encoding(resp.content, is_html=True)
    encoding = html_encoding or http_encoding
    soup = BeautifulSoup(resp.content, 'html.parser', from_encoding=encoding)
    #print("encoding: ", encoding)
    return soup

def special_str_to_int(str):
    a = str.split(",")
    interval = len(a)

    result = 0
    count = 0
    for i in range(interval-1, -1, -1):
        result += int(a[i]) * pow(10, count*3)
        count +=1
    return result

#############################################
#              Parsing stock price          #
#############################################
def twdate(date):
    year  = date.year-1911
    month = date.month
    day   = date.day
    twday = '{}/{:02}/{:02}'.format(year,month,day)
    return twday
   
def downloadOTC(date):      
    url='http://www.tpex.org.tw/web/stock/aftertrading/otc_quotes_no1430/stk_wn1430_print.php?l=zh-tw&d='+twdate(date)+'&se=EW'
    print(url)
    table = pd.read_html(url)[0]
    rowCount = table.values.shape[0]-1
    return table.values[:rowCount]

def downloadHistock(url="https://histock.tw/stock/rank.aspx?m=0&d=0&p=all"):
    url = "https://histock.tw/stock/rank.aspx?m=0&d=0&p=all" if url is None else url
    print(url)
    soup = get_soup(url)
    table = soup.find('table', attrs={'class':'gvTB'})
    rows = table.find_all('tr')
    title = [[x.text for x in rows[0].find_all('div', class_="")][y] for y in [0,1,2,11]]
    return rows, title

#############################################
#              upload excel file            #
#############################################
def authenticate(cred_name):
    SCOPES = ['https://www.googleapis.com/auth/drive.file']
    creds = None

    # load exist certificate
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    # create new certificate 
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            #flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            flow = InstalledAppFlow.from_client_secrets_file(f'{cred_name}.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        # save certificate
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    return creds

def upload_to_drive_as_google_sheet(file_path):
    creds = authenticate(cred_name='client_secret_26780629440-mevsthtu7l2olgs05jb0jkpk8jg7r38p.apps.googleusercontent.com')
    service = build('drive', 'v3', credentials=creds)

    # setup file is google sheet format
    mydate = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_metadata = {
        'name': os.path.basename(file_path).replace('.xlsx', f'_{mydate}'),
        'parents': [folder_id],
        'mimeType': 'application/vnd.google-apps.spreadsheet'
    }
    
    # upload xlsx and convert to google sheet
    media = MediaFileUpload(file_path, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    print(f"File ID: {file.get('id')} uploaded and converted to Google Sheets successfully.")
    print(f"Upload filename: {os.path.basename(file_path).replace('.xlsx', f'_{mydate}')}")

if __name__ == '__main__':
    exec_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    data = []

    # Collect data
    rows, title = downloadHistock()
    for row in rows[1:]:
        item = [x.text.strip() for x in row.find_all('td')]
        a = str(item[0])
        b = str(item[1])
        c = float(item[2])
        d = special_str_to_int(item[11])
        data.append([a,b,c,d])

    # paring OTC stock
    resultOTC = downloadOTC(date.today())
    for item in resultOTC:
        a = str(item[0])
        b = str(item[1])
        c = float(item[2]) if isfloat(item[2]) else '----'
        d = int(item[7])/1000
        data.append([a,b,c,d])

    creater_success = False
    df = pd.DataFrame(data, columns=title)
    df = df.drop_duplicates(subset='代號')
    with pd.ExcelWriter("股價N.xlsx", engine='xlsxwriter') as writer:
        d2 = pd.DataFrame([[str(exec_time), '', ''], title], columns=title)
        df = pd.concat([d2, df[:]]).reset_index(drop=True)
        df.to_excel(writer, sheet_name='股價N', index=False, header=None)

        # set up format
        workbook  = writer.book
        worksheet = writer.sheets['股價N']
        format1 = workbook.add_format({'num_format': '#,##0.00'})
        worksheet.set_column(2, 2, None, format1)  # (from col, to col, cell width, FORMAT)
        creater_success = True

    if creater_success:
        folder_id = '1A6tafYN17dNqV6lkmVG-qBgmCZBAt-U5'
        excel_file_path = '股價N.xlsx'
        upload_to_drive_as_google_sheet(excel_file_path)
    else:
        print(f"Not upload to google drive...")
    


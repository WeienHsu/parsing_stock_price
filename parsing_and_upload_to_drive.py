import os, sys, shutil, math
import urllib3
os.environ['REQUESTS_CA_BUNDLE'] = os.path.join(os.path.dirname(sys.argv[0]), 'cacert.pem')
import pandas as pd
import yfinance as yf
import requests
from time import time
from io import StringIO
from bs4 import BeautifulSoup
from bs4.dammit import EncodingDetector
from datetime import datetime, date, timedelta
# for upload excel to google drive
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

def load_env_variables(env_filename='.env', base_path='./'):
    env_vars = {
        'FOLDER_ID': '', 'CRED_NAME': '', 'FILE_ID': '', 'CP_EXCEL_TO': '',
        'FMP_API_KEY': '', 'CUSTOM_US_STOCKS': ''
    }

    env_path = os.path.join(base_path, env_filename)
    print(f"get .env from current path: {env_path}")
    try:
        with open(env_path, 'r') as file:
            for line in file:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    key, value = line.split('=', 1)
                    env_vars[key] = value.strip()
    except FileNotFoundError:
        print(f"[WARNING] .env file not found at {env_path}. Upload process skipped.")

    return env_vars


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
    #url='http://www.tpex.org.tw/web/stock/aftertrading/otc_quotes_no1430/stk_wn1430_print.php?l=zh-tw&d='+twdate(date)+'&se=EW'
    #url='https://wwwov.tpex.org.tw/web/stock/aftertrading/daily_close_quotes/stk_quote.php?l=zh-tw'
    url='https://www.tpex.org.tw/zh-tw/mainboard/trading/info/post.html'
    print(url)
    table = pd.read_html(url)[0]
    rowCount = table.values.shape[0]-1
    return table.values[:rowCount]

def downloadByCSVUrl(url):
    print(url)

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Language': 'zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7',
        'Referer': 'https://www.twse.com.tw/',
        'Connection': 'keep-alive'
    }

    data = []
    try:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        session = requests.Session()
        response = session.get(url, headers=headers, timeout=30, verify=False)
        response.encoding = 'utf-8'
        if response.status_code == 200:
            csv_content = response.content.decode('utf-8')
            df = pd.read_csv(StringIO(csv_content))
            for index, row in df.iterrows():
                _, sym, stock_name, price, _ = row.to_list()
                data.append([sym,stock_name,price,math.nan])
        else:
            print(f'Download failed..., {response.status_code}')
    except requests.exceptions.RequestException as e:
        raise requests.exceptions.RequestException(f'{e}')

    print(f'{url} ....ok!')
    return data

def downloadByCSVUrl_tpex(url):
    now = datetime.now()
    today = now.date()
    if now.hour < 15:
        target_date = today - timedelta(days=1)  # 往前一天
    else:
        target_date = today  # 当天日期
    #formatted_date = "2024%2F10%2F28"
    formatted_date = target_date.strftime('%Y/%m/%d').replace('/', '%2F')
    url = url.format(formatted_date)

    print(url)
    headers = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36'
    }
    # response = requests.get(url, headers=headers)
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    session = requests.Session()
    response = session.get(url, headers=headers, timeout=30, verify=False)
    data = []
    if response.status_code == 200:
        df = pd.read_csv(StringIO(response.text), sep=',', quotechar='"', skipinitialspace=True, skiprows=2)
        if len(df) <= 2:
            print(f'{url} ..... Not found any data')
            return data
        for index, row in df.iterrows():
            res = row.to_list()
            sym = res[0]
            stock_name = res[1]
            #print(sym,stock_name,res[6])
            #price = float(res[6])
            price = float(str(res[6]).replace(',',''))
            if price == 0.0: continue
            #sym,stock_name,price,_ = row.to_list()
            data.append([sym,str(stock_name).strip(),price,math.nan])
    else:
        print(f'Download failed..., {response.status_code}')
    print(f'{url} ....ok!')
    return data

def downloadHistock(url="https://histock.tw/stock/rank.aspx?m=0&d=0&p=all"):
    url = "https://histock.tw/stock/rank.aspx?m=0&d=0&p=all" if url is None else url
    print(url)
    soup = get_soup(url)
    table = soup.find('table', attrs={'class':'gvTB'})
    rows = table.find_all('tr')
    title = [[x.text for x in rows[0].find_all('div', class_=None)][y] for y in [0,1,2,11]]
    print(f'{url} ....ok!')
    return rows, title

#############################################
#              Parsing american stock       #
#############################################
def get_or_create_sp500_list(csv_filename='stock_list.csv'):
    """
    檢查本地是否有美股名單 CSV：
    - 若有，直接讀取並回傳 DataFrame。
    - 若無，從維基百科爬取 S&P 500 名單，存成本地 CSV，再回傳 DataFrame。
    """
    # 判斷執行環境取得正確路徑 (和原本腳本的寫法一致)
    base_path = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
    csv_path = os.path.join(base_path, csv_filename)

    if os.path.exists(csv_path):
        print(f"找到本地美股名單: {csv_path}，直接讀取。")
        try:
            df = pd.read_csv(csv_path)
            return df
        except Exception as e:
            print(f"讀取本地 CSV 失敗，將重新下載: {e}")
            # 如果讀取失敗，就繼續往下執行重新下載的邏輯
            pass

    print("未找到本地美股名單，正在從維基百科下載 S&P 500 最新成分股...")
    wiki_url = "https://en.wikipedia.org/wiki/List_of_S%26P_500_companies"
    try:
        # Use requests with User-Agent to avoid 403 Forbidden
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        import requests
        response = requests.get(wiki_url, headers=headers)
        response.raise_for_status()

        # read_html 會回傳所有表格，標普 500 名單是第一個 [0]
        tables = pd.read_html(response.text)
        sp500_df = tables[0]

        # 我們只需要 股票代號(Symbol) 和 公司名稱(Security)
        # 為了跟台股的 DataFrame 命名盡量保持一致，我們重命名欄位
        clean_df = sp500_df[['Symbol', 'Security']].copy()
        clean_df.columns = ['代號', '名稱']

        # ⚠️ 注意: 稍微修整一下代號，符合 Yahoo Finance 格式。
        # 維基百科會把 Berkshire 寫成 "BRK.B"，YF 要的是 "BRK-B"
        clean_df['代號'] = clean_df['代號'].astype(str).str.replace('.', '-')

        # 存成 CSV 檔案 (不保留 index)
        clean_df.to_csv(csv_path, index=False, encoding='utf-8-sig')
        print(f"成功下載並建立 {csv_path}，共 {len(clean_df)} 檔股票。")

        return clean_df

    except Exception as e:
        print(f"下載 S&P 500 清單嚴重失敗: {e}")
        return pd.DataFrame()

def download_yfinance_from_csv(csv_filename='stock_list.csv'):
    """
    1. 呼叫上方函式取得股票清單 (從本地或 Wikipedia)
    2. 若有設定 .env 的 CUSTOM_US_STOCKS，也一併加入清單
    3. 使用 yfinance 一次性下載這些股票的今日即時報價
    """
    stock_df = get_or_create_sp500_list(csv_filename)
    data = []

    # 讀取 .env 的自訂美股名單
    base_path = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
    params = load_env_variables(base_path=base_path)
    custom_stocks_str = params.get('CUSTOM_US_STOCKS', '')
    custom_stocks_list = [s.strip() for s in custom_stocks_str.split(',')] if custom_stocks_str else []

    if stock_df.empty and not custom_stocks_list:
        print("無法取得股票清單且無自訂名單，略過美股抓取。")
        return data

    # 將代號與名稱轉為 dictionary 方便最後查找
    symbol_map = {}
    all_symbols = []

    if not stock_df.empty:
        symbol_map = dict(zip(stock_df['代號'], stock_df['名稱']))
        all_symbols = stock_df['代號'].tolist()

    # 將自訂名單加入 (若同名會被集合掉，但這裡簡單 list append 即可)
    for sym in custom_stocks_list:
        if sym and sym not in all_symbols:
            all_symbols.append(sym)
            symbol_map[sym] = sym # 自訂代號沒有中文名稱，就拿代號當名稱顯示

    print(f"開始透過 yfinance 批次抓取 {len(all_symbols)} 檔美股即時報價...")

    # 為了避免 SQLite database is locked 錯誤以及 yfinance 下載大量檔案時發生 TypeError (Yahoo API 阻擋/解析錯誤)
    # 我們將所有股票代號進行「分批次下載」(例如每批次 100 檔)，這樣可以大幅減輕 SQLite 併發寫入的壓力。
    chunk_size = 10
    for i in range(0, len(all_symbols), chunk_size):
        chunk_symbols = all_symbols[i:i + chunk_size]
        try:
            tickers_string = " ".join(chunk_symbols)
            # 關閉多執行緒 (threads=False)，雖然下載速度會稍微慢一點，但可以保證不會因為併發寫入而觸發 SQLite locked 或 API 限流
            hist = yf.download(tickers_string, period="1d", group_by='ticker', threads=False, progress=False)

            if hist.empty:
                print(f"批次 {i+1}~{i+len(chunk_symbols)} 回傳空資料！(可能遇到 API 阻擋或無網路)")
                continue

            for sym in chunk_symbols:
                try:
                    # 若只有抓一檔股票，yfinance 的回傳結構會不一樣
                    if len(chunk_symbols) == 1:
                        latest_data = hist.iloc[-1]
                    else:
                        # 取出該代號的最新一筆資料
                        if sym in hist.columns.levels[0]:
                            latest_data = hist[sym].iloc[-1]
                        else:
                            continue

                    price = float(latest_data['Close'])
                    volume = float(latest_data['Volume'])

                    # 過濾明顯異常沒有價格的資料
                    if pd.isna(price) or price <= 0:
                        continue

                    # 從字典拿取正確的公司名稱
                    stock_name = symbol_map.get(sym, sym)

                    # 放入你的統一 array 中 [代號, 名稱, 價格, 成交量]
                    data.append([sym, stock_name, price, volume])

                except Exception as e:
                    # 即使某一檔出了小問題，靜默跳過，不影響其他
                    continue

        except Exception as e:
            print(f"yfinance 批次下載報價發生嚴重錯誤 ({i+1}~{i+len(chunk_symbols)}): {e}")

    print(f"美股報價抓取完成！成功取得 {len(data)} 檔報價。")
    return data

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
            print(creds, creds.expired, creds.refresh_token)
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(f'{cred_name}.json', SCOPES)
            creds = flow.run_local_server(port=0)

        # save certificate
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    return creds


def upload_to_drive_as_google_sheet(file_path, cred_name, folder_id, file_id=None):
    creds = authenticate(cred_name=cred_name)
    service = build('drive', 'v3', credentials=creds)

    # 設定檔案基本名稱（不包含日期時間戳）
    base_filename = os.path.basename(file_path).replace('.xlsx', '')
    if file_id:
        # 如果提供了file_id，則更新現有檔案
        media = MediaFileUpload(file_path, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        file = service.files().update(
            fileId=file_id,
            media_body=media,
            fields='id'
        ).execute()
        print(f"File ID: {file.get('id')} updated successfully.")
        print(f"Updated filename: {base_filename}")
        return file.get('id')
    else:
        # 如果沒有提供file_id，則檢查是否存在同名檔案
        query = f"name='{base_filename}' and '{folder_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
        results = service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
        items = results.get('files', [])
        if items:
            # 如果找到同名檔案，則更新它
            file_id = items[0]['id']
            media = MediaFileUpload(file_path, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            file = service.files().update(
                fileId=file_id,
                media_body=media,
                fields='id'
            ).execute()
            print(f"File ID: {file.get('id')} updated successfully.")
            print(f"Updated filename: {base_filename}")
            return file.get('id')
        else:
            # 如果沒有找到同名檔案，則創建新檔案
            file_metadata = {
                'name': base_filename,
                'parents': [folder_id],
                'mimeType': 'application/vnd.google-apps.spreadsheet'
            }
            media = MediaFileUpload(file_path, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            file = service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id'
            ).execute()
            print(f"File ID: {file.get('id')} uploaded and converted to Google Sheets successfully.")
            print(f"Upload filename: {base_filename}")
            return file.get('id')

if __name__ == '__main__':
    exec_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    data_tw = []

    # ======== 1. 收集台股資料 ========
    tw_time = time()
    rows, title = downloadHistock()
    for row in rows[1:]:
        item = [x.text.strip() for x in row.find_all('td')]
        a = str(item[0])  # 代號
        b = str(item[1])  # 名稱
        c = float(item[2])  # price
        d = special_str_to_int(item[11])  # 成交量
        data_tw.append([a,b,c,d])

    resultOTC = downloadByCSVUrl(url='https://www.twse.com.tw/exchangeReport/STOCK_DAY_AVG_ALL?response=open_data')
    data_tw.extend(resultOTC)

    resultTPEX = downloadByCSVUrl_tpex(url="https://www.tpex.org.tw/www/zh-tw/afterTrading/fixPricing?date={}&id=&response=csv")
    data_tw.extend(resultTPEX)
    print(f" ==== Taiwan Total count: {len(data_tw)} (cost time: {time()-tw_time:.2f}s)===")

    # 將台股轉為 DataFrame 並去重
    df_tw = pd.DataFrame(data_tw, columns=title)
    df_tw = df_tw.drop_duplicates(subset='代號').reset_index(drop=True)

    # ======== 2. 收集美股資料 ========
    us_time = time()
    # 呼叫美股抓取函式
    resultUS = download_yfinance_from_csv('stock_list.csv')
    print(f" ==== US Total count: {len(resultUS)} (cost time: {time()-us_time:.2f}s)====")

    # 將美股轉為 DataFrame
    title_us = ['美股代號', '美股名稱', '美股價格', '美股成交量']
    df_us = pd.DataFrame(resultUS, columns=title_us)
    df_us = df_us.drop_duplicates(subset='美股代號').reset_index(drop=True)

    # ======== 3. 合併台美股 (左右並排) ========
    # 使用 axis=1 進行水平合併
    df_combined = pd.concat([df_tw, df_us], axis=1)

    # ======== 4. 存檔到 Excel ========
    creater_success = False
    if getattr(sys, 'frozen', False):  # 如果是打包后的应用
        current_directory = os.path.dirname(sys.executable)  # 获取执行文件的目录
    else:
        current_directory = os.path.dirname(os.path.abspath(__file__))  # 获取脚本的目录
    xlsx_path = os.path.join(current_directory, "股價N.xlsx")
    print(f"File Path: {xlsx_path}")

    with pd.ExcelWriter(xlsx_path, engine='xlsxwriter') as writer:
        # 製作特殊的置頂表頭 (包含兩邊的欄位)
        combined_titles = title + title_us
        d2 = pd.DataFrame([[str(exec_time)] + [''] * (len(combined_titles) - 1), combined_titles], columns=combined_titles)

        # 覆蓋回 df_combined 的欄位名稱，讓它可以被安全 concat
        df_combined.columns = combined_titles
        df_final = pd.concat([d2, df_combined]).reset_index(drop=True)

        df_final.to_excel(writer, sheet_name='股價N', index=False, header=None)

        # set up format
        workbook  = writer.book
        worksheet = writer.sheets['股價N']
        format_price = workbook.add_format({'num_format': '#,##0.00'})

        # 第 C 欄 (台股價格) 套用格式
        worksheet.set_column(2, 2, None, format_price)
        # 第 G 欄 (美股價格) 套用格式 (索引為 6)
        worksheet.set_column(6, 6, None, format_price)

        creater_success = True

    # get parameter for using drive api if it's exist
    base_path = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
    params = load_env_variables(base_path=base_path)
    folder_id = params['FOLDER_ID']
    cred_name = os.path.join(base_path, params['CRED_NAME'])
    if creater_success and folder_id and cred_name:
        upload_to_drive_as_google_sheet(
            xlsx_path, cred_name=cred_name, folder_id=folder_id)
    else:
        print(f"Not upload to google drive...")

    cp_path = params['CP_EXCEL_TO']
    if cp_path and os.path.exists(cp_path):
        try:
            shutil.copy2(xlsx_path, cp_path)
            print(f"Copy to {cp_path}, Done!")
        except:
            print(f"Copy file failed!")

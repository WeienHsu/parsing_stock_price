import os, sys
os.environ['REQUESTS_CA_BUNDLE'] = os.path.join(os.path.dirname(sys.argv[0]), 'cacert.pem')
import pandas as pd
import requests
import xlsxwriter
# from urllib.request import urlopen
# from urllib.request import Request
from bs4 import BeautifulSoup
from bs4.dammit import EncodingDetector
from datetime import datetime

#print(os.getcwd())

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


if __name__ == '__main__':
    exec_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    myurl = "https://stock.wespai.com/lists"
    soup = get_soup(myurl)
    # raw_request = Request(myurl)
    # raw_request.add_header('User-Agent', 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:78.0) Gecko/20100101 Firefox/78.0')
    # raw_request.add_header('Accept', 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8')
    # resp = urlopen(raw_request).read()
    # soup = BeautifulSoup(resp, 'html.parser', from_encoding='UTF-8')
    # # print(soup.prettify())
    # # print(soup)

    table = soup.find('table', attrs={'class':'display'})

    # parsing tabel title
    title = []
    table_title = table.find('thead')
    names = table_title.find_all('tr')
    for name in names:
        cols = name.find_all('th')
        # title = [ele.text.strip() for ele in cols]
        # title.append([ele.text.strip() for ele in cols])

        title = [cols[x].text.strip() for x in [0,1,3]]


    # parsing table content
    data = []
    table_body = table.find('tbody')
    rows = table_body.find_all('tr')
    for row in rows:
        cols = row.find_all('td')
        # cols = [ele.text.strip() for ele in cols]
        # data.append(cols)
        # data.append([str(cols[x].text.strip()) for x in [0,1,3]])
        a = str(cols[0].text.strip())
        b = str(cols[1].text.strip())
        c = round(float(cols[3].text.strip()),2) if isfloat(cols[3].text.strip()) else str(cols[3].text.strip())
        data.append([a,b,c])

    # convert to excel
    df = pd.DataFrame(data, columns=title)
    with pd.ExcelWriter("股價N.xlsx", engine='xlsxwriter') as writer:
        d2 = pd.DataFrame([[str(exec_time), '', ''], title], columns=title)
        df = pd.concat([d2, df[:]]).reset_index(drop=True)
        df.to_excel(writer, sheet_name='股價N', index=False, header=None)

        # set up format
        workbook  = writer.book
        worksheet = writer.sheets['股價N']
        format1 = workbook.add_format({'num_format': '#,##0.00'})
        worksheet.set_column(2, 2, None, format1)  # (from col, to col, cell width, FORMAT)
        writer.save()

    #df.loc[-1] = [str(exec_time), '', '']
    #d2 = pd.DataFrame([title, [str(exec_time), '', '']], columns=title)
    # d2 = pd.DataFrame([[str(exec_time), '', ''], title], columns=title)
    # df = pd.concat([d2, df[:]]).reset_index(drop=True)
    # df.to_excel('股價N.xls', sheet_name='股價N', index=False, header=None)
    
    



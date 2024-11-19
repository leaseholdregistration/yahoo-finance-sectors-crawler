import requests
import re
import os
from bs4 import BeautifulSoup
from datetime import datetime
from openpyxl import Workbook

sectors = [
    "technology",
    "financial-services",
    "healthcare",
    "consumer-cyclical",
    "communication-services",
    "industrials",
    "consumer-defensive",
    "energy",
    "basic-materials",
    "real-estate",
    "utilities"
]

def fetch_data():
    base_url = "https://finance.yahoo.com/sectors/{}"
    all_data = {sector: [] for sector in sectors}
    headers = {"Accept-Language": "en-US,en;q=0.9"}

    for sector in sectors:
        url = base_url.format(sector)
        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()
        except requests.RequestException as e:
            print(f"요청 오류: {e}")
            return

        soup = BeautifulSoup(response.content, "html.parser")
        heatmap_container = soup.find('div', class_='heatMap-container')
        if not heatmap_container:
            print(f"{sector} 페이지에서 heatMap-container를 찾을 수 없습니다.")
            continue

        # 패턴에 맞는 a 태그 찾기
        pattern = re.compile(r'none-link.*fin-size-medium|fin-size-medium.*none-link')
        sector_links = heatmap_container.find_all('a', class_=pattern)

        if not sector_links:
            print(f"{sector} 페이지에서 데이터를 찾을 수 없습니다.")
            continue

        for link in sector_links:
            sector_name = link.find('div', class_='ticker-div').text.strip()
            percent_change = link.find('div', class_='percent-div').text.strip()
            all_data[sector].append((sector_name, percent_change))

        # 데이터 정렬: sector_name 기준으로 정렬
        all_data[sector].sort(key=lambda x: x[0])

    if not any(all_data.values()):
        print("데이터를 찾을 수 없습니다.")
        return

    # 엑셀 파일로 저장
    wb = Workbook()
    ws = wb.active

    # 열 헤더 추가
    headers = ["page"] + [sector for sector in sectors for _ in all_data[sector]]
    ws.append(headers)

    # sector 행 추가
    sector_row = ["sector"] + [sector_name for sector in sectors for sector_name, _ in all_data[sector]]
    ws.append(sector_row)

    # change 행 추가
    change_row = ["change"] + [change for sector in sectors for _, change in all_data[sector]]
    ws.append(change_row)

    today_date = datetime.today().strftime('%Y-%m-%d')
    file_name = f"output_{today_date}.xlsx"
    wb.save(file_name)

    print(f"데이터를 엑셀 파일로 저장했습니다: {file_name}")
    return file_name

def send_email(file_path):
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.application import MIMEApplication

    EMAIL_ADDRESS = os.environ['EMAIL_ADDRESS']
    EMAIL_PASSWORD = os.environ['EMAIL_PASSWORD']
    recipient = os.environ['RECIPIENT_EMAIL']

    msg = MIMEMultipart()
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = recipient
    msg['Subject'] = '크롤링 데이터 결과'

    # 엑셀 파일 첨부
    with open(file_path, 'rb') as file:
        attachment = MIMEApplication(file.read(), _subtype="xlsx")
        attachment.add_header('Content-Disposition', 'attachment', filename=file_path)
        msg.attach(attachment)

    # 이메일 발송
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    server.send_message(msg)
    server.quit()

    print("이메일이 발송되었습니다.")

if __name__ == "__main__":
    excel_file = fetch_data()
    if excel_file:
        send_email(excel_file)

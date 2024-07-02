import os
import tkinter as tk
from tkinter import messagebox, ttk
import requests
from bs4 import BeautifulSoup
from datetime import datetime
import pytz
import webbrowser
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

    # 로딩바 초기화
    progress_bar["value"] = 0
    progress_bar["maximum"] = len(sectors)
    button_fetch.config(state=tk.DISABLED)

    for sector in sectors:
        url = base_url.format(sector)
        try:
            response = requests.get(url)
            response.raise_for_status()
        except requests.RequestException as e:
            messagebox.showerror("요청 오류", f"웹사이트에 접근할 수 없습니다. 오류: {e}")
            button_fetch.config(state=tk.NORMAL)
            return

        soup = BeautifulSoup(response.content, "html.parser")
        heatmap_container = soup.find('div', class_='heatMap-container')
        if not heatmap_container:
            messagebox.showinfo("데이터 없음", f"{sector} 페이지에서 heatMap-container를 찾을 수 없습니다.")
            progress_bar["value"] += 1
            root.update_idletasks()
            continue

        sector_links = heatmap_container.find_all('a', class_='none-link fin-size-medium svelte-wdkn18')

        if not sector_links:
            messagebox.showinfo("데이터 없음", f"{sector} 페이지에서 데이터를 찾을 수 없습니다.")
            progress_bar["value"] += 1
            root.update_idletasks()
            continue

        for link in sector_links:
            sector_name = link.find('div', class_='ticker-div').text.strip()
            percent_change = link.find('div', class_='percent-div').text.strip()
            all_data[sector].append((sector_name, percent_change))

        # 데이터 정렬: sector_name 기준으로 정렬
        all_data[sector].sort(key=lambda x: x[0])

        # 로딩바 업데이트
        progress_bar["value"] += 1
        root.update_idletasks()

    if not any(all_data.values()):
        messagebox.showinfo("데이터 없음", "데이터를 찾을 수 없습니다.")
        button_fetch.config(state=tk.NORMAL)
        return

    # openpyxl을 사용하여 데이터 엑셀 파일로 저장
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

    messagebox.showinfo("완료", f"데이터를 엑셀 파일로 저장했습니다: {file_name}")
    progress_bar["value"] = 0
    button_fetch.config(state=tk.NORMAL)


def update_time_and_status():
    # 미국 동부 시간 (ET) 설정
    eastern = pytz.timezone('US/Eastern')
    current_time = datetime.now(eastern)
    current_time_str = current_time.strftime('%Y-%m-%d %H:%M:%S')

    # 한국 시간 (KST)
    korea = pytz.timezone('Asia/Seoul')
    current_time_kst = current_time.astimezone(korea)
    current_time_kst_str = current_time_kst.strftime('%Y-%m-%d %H:%M:%S')

    # 주식 시장 운영 시간 확인 (월-금, 09:30-16:00 EST)
    market_open = current_time.replace(hour=9, minute=30, second=0, microsecond=0)
    market_close = current_time.replace(hour=16, minute=0, second=0, microsecond=0)
    market_status = "Closed"

    if current_time.weekday() < 5 and market_open <= current_time <= market_close:
        market_status = "Open"

    time_label.config(text=f"현 미국 동부 시간: {current_time_str}\n현 한국 시간: {current_time_kst_str}")
    status_label.config(text=f"증권 거래 개장 여부: {market_status}")
    root.after(1000, update_time_and_status)


def open_url(event):
    webbrowser.open_new("https://finance.yahoo.com/sectors")

def open_folder():
    os.startfile(os.getcwd())

# GUI 설정
root = tk.Tk()
root.title("Yahoo Finance Sectors Crawler")

frame = tk.Frame(root)
frame.pack(pady=20)

# 미국 시간 및 주식 시장 상태 표시 라벨
time_label = tk.Label(frame, text="")
time_label.grid(row=0, column=0, columnspan=2, padx=10, pady=5)

status_label = tk.Label(frame, text="")
status_label.grid(row=1, column=0, columnspan=2, padx=10, pady=5)

# 데이터 가져오기 버튼
button_fetch = tk.Button(frame, text="데이터 가져오기", command=fetch_data)
button_fetch.grid(row=2, column=0, columnspan=2, padx=10, pady=10)

# 저장 폴더 열기 버튼
button_open_folder = tk.Button(frame, text="저장 폴더 열기", command=open_folder)
button_open_folder.grid(row=3, column=0, columnspan=2, padx=10, pady=10)

# 로딩바
progress_bar = ttk.Progressbar(frame, orient="horizontal", length=300, mode="determinate")
progress_bar.grid(row=4, column=0, columnspan=2, pady=10)

# Yahoo Finance 링크
link = tk.Label(frame, text="Go to Yahoo Finance Sectors", fg="blue", cursor="hand2")
link.grid(row=5, column=0, columnspan=2, pady=10)
link.bind("<Button-1>", open_url)

update_time_and_status()
root.mainloop()

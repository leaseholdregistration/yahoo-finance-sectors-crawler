import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
import pytz

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
    all_data = []

    # 로딩바 초기화
    progress_bar["value"] = 0
    progress_bar["maximum"] = len(sectors)

    for sector in sectors:
        url = base_url.format(sector)
        try:
            response = requests.get(url)
            response.raise_for_status()
        except requests.RequestException as e:
            messagebox.showerror("요청 오류", f"웹사이트에 접근할 수 없습니다. 오류: {e}")
            return

        soup = BeautifulSoup(response.content, "html.parser")
        heatmap_container = soup.find('div', class_='heatMap-container')
        if not heatmap_container:
            messagebox.showinfo("데이터 없음", f"{sector} 페이지에서 heatMap-container를 찾을 수 없습니다.")
            continue

        sector_links = heatmap_container.find_all('a', class_='none-link fin-size-medium svelte-wdkn18')

        if not sector_links:
            messagebox.showinfo("데이터 없음", f"{sector} 페이지에서 데이터를 찾을 수 없습니다.")
            continue

        for link in sector_links:
            sector_name = link.find('div', class_='ticker-div').text.strip()
            percent_change = link.find('div', class_='percent-div').text.strip()
            all_data.append({'page': sector, 'sector': sector_name, 'change': percent_change})

        # 로딩바 업데이트
        progress_bar["value"] += 1
        root.update_idletasks()

    if not all_data:
        messagebox.showinfo("데이터 없음", "데이터를 찾을 수 없습니다.")
        return

    df = pd.DataFrame(all_data)

    today_date = datetime.today().strftime('%Y-%m-%d')
    file_name = f"output_{today_date}.xlsx"

    df = df.T  # Transpose the DataFrame
    df.to_excel(file_name, index=False, header=False)

    messagebox.showinfo("완료", f"데이터를 엑셀 파일로 저장했습니다: {file_name}")


def update_time_and_status():
    # 미국 동부 시간 (ET) 설정
    eastern = pytz.timezone('US/Eastern')
    current_time = datetime.now(eastern)
    current_time_str = current_time.strftime('%Y-%m-%d %H:%M:%S')

    # 주식 시장 운영 시간 확인 (월-금, 09:30-16:00)
    market_open = current_time.replace(hour=9, minute=30, second=0, microsecond=0)
    market_close = current_time.replace(hour=16, minute=0, second=0, microsecond=0)
    market_status = "Closed"

    if current_time.weekday() < 5 and market_open <= current_time <= market_close:
        market_status = "Open"

    time_label.config(text=f"Current US Eastern Time: {current_time_str}")
    status_label.config(text=f"Market Status: {market_status}")
    root.after(1000, update_time_and_status)


# GUI 설정
root = tk.Tk()
root.title("program")

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

# 로딩바 추가
progress_bar = ttk.Progressbar(frame, orient="horizontal", length=300, mode="determinate")
progress_bar.grid(row=3, column=0, columnspan=2, pady=10)

update_time_and_status()
root.mainloop()
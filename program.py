import yfinance as yf
import pandas as pd
from tkinter import Tk, Listbox, Button, Entry, Label, Radiobutton, StringVar, messagebox
from datetime import datetime

def get_stock_data(tickers):
    data = []
    for ticker in tickers:
        stock = yf.Ticker(ticker)
        info = stock.info
        data.append([
            info['symbol'],
            info['longName'],
            info['exchange'],
            info['country'],
            info['marketCap'],
            info['sharesOutstanding'],
            info['regularMarketPrice'],
            info['trailingPE'],
            info['forwardPE'],
            info['enterpriseToEbitda'],
            info['returnOnEquity'],
            info['priceToBook'],
            info['earningsQuarterlyGrowth'],
            info['sector'],
            info['trailingAnnualDividendYield'],
            info['regularMarketPrice']  # Assume this as price growth (YoY)
        ])
    return pd.DataFrame(data, columns=[
        'Ticker', 'Name', 'Exchange', 'Country', 'Market Cap',
        'Shares Outstanding', 'Price per Share', 'Trailing P/E', 'Forward P/E',
        'EV/EBITDA', 'ROE', 'PBR', 'EPS Growth', 'Sector', 'Dividend Yield', 'Price Growth (YoY)'
    ])

def export_to_excel(data):
    current_datetime = datetime.now().strftime("%Y%m%d%H%M%S")
    filename = f'output-{current_datetime}.xlsx'
    data.to_excel(filename, index=False)
    return filename

def on_export():
    selected_markets = [listbox.get(idx) for idx in listbox.curselection()]
    tickers = []  # You would need to obtain a list of tickers for the selected markets
    data = get_stock_data(tickers)

    # 필터링
    if per_value.get():
        filter_value = float(per_value.get())
        filter_direction = per_direction.get()
        if filter_direction == 'Below':
            data = data[data['Trailing P/E'] <= filter_value]
        else:
            data = data[data['Trailing P/E'] >= filter_value]

    # ... 같은 방식으로 다른 지표들에 대해 필터링 ...

    filename = export_to_excel(data)
    messagebox.showinfo("Export Complete", f"Data exported to {filename}")

root = Tk()
root.title("program")

listbox = Listbox(root, selectmode='extended')  # 복수 선택을 위해 'extended'로 설정
listbox.insert('end', 'NASDAQ')
listbox.insert('end', 'KOSDAQ')
# ... add other markets
listbox.pack()

# 지표 필터링 입력
per_label = Label(root, text="P/E Ratio")
per_label.pack()
per_value = Entry(root)
per_value.pack()
per_direction = StringVar(value='Below')
per_below = Radiobutton(root, text="Below", variable=per_direction, value='Below')
per_below.pack()
per_above = Radiobutton(root, text="Above", variable=per_direction, value='Above')
per_above.pack()

# ... 같은 방식으로 다른 지표들에 대한 입력 위젯 추가 ...

export_button = Button(root, text="Export to Excel", command=on_export)
export_button.pack()

root.mainloop()

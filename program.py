import yfinance as yf
import pandas as pd
from tkinter import Tk, Listbox, Button, Entry, Label, Radiobutton, StringVar, messagebox
from datetime import datetime

#TICKERS = ['AAPL', 'MSFT']
TICKERS = ['AAPL', 'MSFT', 'AMZN', '035420.KS']
INFO_KEYS = ['symbol', 'shortName', 'exchange', 'country', 'marketCap', 'currency', 'sharesOutstanding',
        'trailingPE', 'forwardPE', 'enterpriseToEbitda', 'returnOnEquity', 'priceToBook',
        'earningsQuarterlyGrowth', 'sector', 'trailingAnnualDividendYield']

def get_stock_data(tickers):
    data = []
    for ticker in tickers:
        row_data = []
        stock = yf.Ticker(ticker)
        info = stock.info
        for key in INFO_KEYS:
            try:
                row_data.append(info[key])
            except Exception as e:
                print(f"caught except: {ticker} - {e}")
                row_data.append('N/A')
        data.append(row_data)
    return pd.DataFrame(data, columns=[
        '티커', '종목명', '거래소', '국가', '시가총액', '화폐',
        '유통 주식수', 'trailingPER (직전 12개월)', 'forwardPER (직후 12개월 추정)',
        'EV/EBITDA', 'ROE', 'PBR', 'EPS Growth (분기별 수익 성장률)', '업종', '최근1년간 배당 수익률'
    ])

def export_to_excel(data):
    current_datetime = datetime.now().strftime("%Y%m%d%H%M%S")
    filename = f'output-{current_datetime}.xlsx'
    data.to_excel(filename, index=False)
    return filename

def on_export():
    tickers = TICKERS  # You would need to obtain a list of tickers for the selected markets
    data = get_stock_data(tickers)

    filename = export_to_excel(data)
    messagebox.showinfo("Export Complete", f"Data exported to {filename}")

root = Tk()
root.title("주식 데이터 파서")

export_button = Button(root, text="Export to Excel", command=on_export)
export_button.pack()

root.mainloop()

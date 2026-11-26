"""
TSMC Excel 每日更新腳本 - 2026-07-01
執行：python update_excel_daily.py
"""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import date

EXCEL_PATH = r"C:\Users\K748\OneDrive - 財團法人中華民國對外貿易發展協會\FET\Stock分析\TSMC_股市分析報告.xlsx"

def make_border():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)

def style_cell(cell, bg="FFFFFF", align="center", bold=False, color="000000"):
    cell.font = Font(name="Arial", size=10, bold=bold, color=color)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    cell.border = make_border()

today_str = '2026-11-26'
tw_price = 'NT$2,150（11/25收，-80、-3.59%；開2,150／高2,170／低2,020，開盤＝收盤的十字線（下影130元）；2026年101個下跌日第8大；跌回MA10 2,176之下、守在MA5 2,102之上；回補2,050–2,230向上缺口、留下2,170–2,230向下缺口；MA60 2,381.25下穿MA120 2,382.25；MACD柱-14.51續負收斂、KD黃金交叉（K 37.92／D 36.51）、RSI 43.08；外資賣超26,971張（重回門檻之外）、投信買超4,791張、自營賣超438張、三大合計-22,618張；放量下跌（字面與實質一致）；雙條件0/2、維持看空、轉弱程度改判加劇）'
change_pct = '-3.59%（11/25收，相對大盤+1.22pp（精算+1.2152），延續相對強勢；加權-1,926.92點、-4.80%收38,195.43，全市場成交12,297.62億）；11/24為+9.85%'
nyse_price = 'US$342.34（11/25收，-1.96%，2年237個下跌日第87大；收在11/05收盤353.10之下3.05%；費半-0.65%；ADR溢價同日口徑+5.50%（ADR 11/25對台股11/25與匯率33.13，台股跌幅較大與新台幣貶值造成擴大），11/24同日口徑+3.16%；盤後約347（約+1.5%，19:59 ET）；11/26美股感恩節休市）；11/24收US$349.20'
volume = '84,912（11/25收，+542.64%；為前報基準55,626張的1.53倍＝放量下跌，剔除11/24鎖漲停極小量的對照基準亦為放量下跌；新5日均量62,450張；2026年216日中由大到小第6）；11/24為13,213張'
news_summary = """11/25低開收十字線2,150(-3.59%)、跌回MA10、留下2,170–2,230向下缺口、MA60下穿MA120；外資賣超26,971張重回門檻之外、投信續買；放量下跌；USTR擬對台課最高15%關稅；美股費半-0.65%、ADR-1.96%（皆在台股收盤後）；維持看空、轉弱改判加劇。"""
change_color = 'FF0000'
try:
    wb = load_workbook(EXCEL_PATH)
    ws = wb["每日更新記錄"]

    # 若最後一行日期 = 今日，更新該行；否則新增一行（避免重複 row）
    last_row = ws.max_row
    last_date = ws.cell(last_row, 1).value
    target_row = last_row if last_date == today_str else last_row + 1
    mode = "更新" if target_row == last_row else "新增"

    row_data = [
        today_str,
        tw_price,
        change_pct,
        nyse_price,
        volume,
        news_summary,
        "Yahoo Finance / Bloomberg"
    ]
    bg = "E9F2FB" if target_row % 2 == 0 else "FFFFFF"
    for col, val in enumerate(row_data, 1):
        c = ws.cell(row=target_row, column=col, value=val)
        if col == 3:
            style_cell(c, bg=bg, align="center", color=change_color, bold=True)
        elif col == 6:
            style_cell(c, bg=bg, align="left")
        else:
            style_cell(c, bg=bg)

    ws["A2"] = f"最後更新：{today_str}"
    wb["封面總覽"]["A3"] = f"報告更新日期：{today_str}"
    wb.save(EXCEL_PATH)
    print(f"Excel {mode}成功：第 {target_row} 行（{today_str}）")
except Exception as e:
    print(f"Excel 更新失敗：{e}")

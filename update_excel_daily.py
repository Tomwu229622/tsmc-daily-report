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

today_str = '2026-09-24'
tw_price = 'NT$2,500（9/23收，7/01以來最高收盤；開盤＝全日最低2,475）'
change_pct = '+1.63%（9/23收，+40.00，領先大盤+0.75%）'
nyse_price = 'US$446.57（9/23收，-1.20%，中止連5漲；美債10年期殖利率升至5.11%）'
volume = '22,818（9/23收，+3.7%，低於5日均量4.9%）'
news_summary = """報告日2026-09-24(Thu)，台股資料為2026-09-23(Wed)官方收盤；ADR、費半與美股個股、韓股為9/23正式收盤；期貨為TAIFEX 9/23日盤，另含今晨夜盤(9/23 15:00至9/24約05:00)。

(1)⭐⭐收復：台股2330開盤即全日最低2,475、收2,500(+40、+1.63%)，為7/01以來最高收盤(2026年收盤≥2,500僅6/22、7/01與本日)；外資轉買7,123張(約NT$178.1億，占全市場外資淨買超NT$373.13億約47.7%)，領先加權+0.75%達0.88pp。

(2)⭐⭐前報六項檢核：價格分界2,445守住且收盤站回2,480之上→9/22黑K改判為連假前單日部位調整；外資轉買→9/22賣超視為單日調節；相對強弱中止連3日落後→輪動判定降級為觀察中；夜盤開盤與收盤方向皆兌現；9/23週結算P/C 79.83%只讀水位。

(3)⭐技術：MACD柱+12.28(9/09以來最大)、RSI 60.97(7/01以來最高)、⚠️J 100.11(8/13以來首度站上100)；五線完整多頭排列連2日、收盤距布林上軌2,501僅1.38元；量22,818張低於5日均量，屬價漲量平。

(4)⚠️⚠️台股收盤後美國10年期公債殖利率升至5.11%(CNN稱19年來首見)，費半-1.23%中止連6紅、ADR-1.20%收$446.57，溢價收斂至+13.18%；今晨夜盤CDF收2,492(-18)、台指期-427點；外資台指期淨空擴大至-76,084口。

(5)立場維持「中性」。次日檢核：收盤能否守2,475(守住則回升中性偏多、跌破2,460則9/23漲幅全數回吐)；9/24為連假前最後交易日，9/25、9/28休市、美股照常交易三日；美光美東9/30財報。"""
change_color = '00B050'
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

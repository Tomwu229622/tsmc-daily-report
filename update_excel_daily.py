"""
TSMC Excel 每日更新腳本 - 2026-05-12
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

today_str = "2026-05-14"
tw_price = "NT$2,225"        # 5/14 Thu 收盤（+15，+0.68% vs 5/13 NT$2,210；跟進 ADR 反彈訊號）
change_pct = "+0.68%"        # 台股 2330 5/14 收盤漲跌幅
nyse_price = "US$399.80"     # NYSE TSM 最新收盤（5/13 Wed，+2.52，+0.63% vs 5/12 $397.28）
volume = "38,800 張"         # 5/14 成交量（反彈量縮整理）
news_summary = "台股2330 5/14反彈收NT$2,225(+15,+0.68% vs 5/13 NT$2,210)、量縮38,800張、市值NT$57.7兆;跟進NYSE TSM 5/13 +0.63%反彈訊號、守NT$2,200心理整數雙重支撐;🏆重大訊息:NVIDIA確認超越Apple為TSMC最大客戶(22% vs 17%、$33B vs $27B年貢獻、$95B採購承諾揭露 vs 兩年前$16B);NYSE TSM 5/13收$399.80(+2.52,+0.63% vs 5/12 $397.28)守$400整數;ADR溢價自+13.9%收斂至+11.8%;5/14估三大法人合計買超+6,500張——外資轉買+5,800張、投信+500張、自營+200張、外資持股回升74.9%;⭐利多續存:(1)NVIDIA確認TSMC最大客戶結構性強化;(2)台灣對美半導體投資將超過$200B+;(3)TSMC Arizona追加$20B資本注入;(4)AMAT $5B EPIC Center AI晶片R&D合作;(5)2nm 2026-2028產能CAGR +70%、NVIDIA為A16首發客戶;(6)SOX YTD +65%、全球半導體2026預估$975B創高;SOX 5/13+0.69%收10,970、NVDA+1.07%、AMD+0.77%、AVGO+0.81%、ASML+0.50%;Barclays$470對應NT$2,450(剩+10.1%)、共識NT$2,320(剩+4.3%);NVDA 5/28 Q1 FY27財報距14天為最重要AI算力指標、TSMC 5月月營收預計6/10前公布;技術面RSI升至51站回中軸、KDJ K(60)/D(58)/J(64)黃金交叉訊號出現、MACD+1.4紅柱暫停翻黑;⚠️Apple轉單Intel 18A傳聞影響部分消化、Intel 5/13回吐-1.39%反映市場質疑;明日5/15關鍵:站回5MA NT$2,235+攻NT$2,255反彈目標"
change_color = "00B050"      # 上漲綠色（+0.68%）

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

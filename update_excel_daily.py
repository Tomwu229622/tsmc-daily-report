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

today_str = "2026-07-07"
tw_price = "NT$2,490（7/7收）"   # 台股2330 7/7 Tue官方收NT$2,490(+30、+1.22% vs 7/6收NT$2,460、量估約24,800張、盤中高觸NT$2,505)為最新確認收盤；連二漲、補漲兌現
change_pct = "+1.22%（7/7收）"    # 台股7/7收盤漲跌幅(+1.22%、連二漲、補漲兌現、跟進NYSE反彈+Citi升評)
nyse_price = "US$451.79（7/6收）" # NYSE TSM 7/6 Mon官方收$451.79(+$17.63、+4.06%、stockanalysis.com確認、前收7/2 $434.16)為最新確認NYSE收盤——假期後重新開盤即V型大反彈、費半SOX同步+3.92%、一舉收復7/1(-6.98%)+7/2(-2.27%)兩日全數跌幅；NYSE 7/7盤待今晚GMT+8
volume = "24,800（7/7收，估）"    # 台股7/7收盤成交量(張、估、上漲日放量)
news_summary = "報告日2026-07-07(Tue)。台股補漲兌現、Citi大幅上調目標至NT$3,800：🚀Citi(花旗)7/6大幅上調TSMC台股目標價自NT$2,875至NT$3,800(+32%、維持買入、自7/7收NT$2,490具+52.6%空間)、為7/16法說會前最新升評、料上修2026營收成長與2027-28資本支出至$75-80B；⚖️台股2330 7/7收NT$2,490(+30、+1.22%、量估約24,800張、盤中高觸NT$2,505逼近NT$2,535史高)連二漲、補漲兌現、跟進NYSE 7/6大反彈+Citi升評雙催化；🚀NYSE TSM 7/6假期後重新開盤即大漲+4.06%(+$17.63)收$451.79(前收7/2 $434.16)、費半SOX同步+3.92%、一舉收復7/1(-6.98%)+7/2(-2.27%)兩日全數跌幅、AI半導體回神、確認為估值修正非基本面惡化；📊台美ADR溢價收斂至約+16.2%(7/6 ADR $451.79折合台股NT$2,893.71 vs 台股7/7 NT$2,490)。🚀分析師續大幅看多：Citi NT$3,800、Barclays維持$625(自$451.79具+38.3%空間)、Nomura NT$3,425(自NT$2,490具+37.6%空間)；⚠️惟Goldman 7/1移出APAC Conviction List(戰術調整非降評)。🤝NVIDIA-SK海力士宣布多年期記憶體合作；NVIDIA-TSMC導入AI入廠提升良率；BofA上修2026全球半導體市場至$1.3兆。⭐漲價續發酵、7nm以下全先進製程調漲5-10%(佔晶圓營收~75%)。中長線基底未變：NVIDIA #1客戶(22%/$33B)/A16首發、Vera Rubin 6顆晶片全TSMC代工、2nm 70-80%良率。⚠️估值警示TSM P/E~39.3x、台股~36.6x；6月營收(7月上旬)與7/16財報為關鍵catalyst；下方支撐NT$2,460/NT$2,445、上方壓力NT$2,505/NT$2,535史高。"
change_color = "00B050"        # 綠色（台股 7/7 收盤上漲 +1.22% 連二漲補漲兌現；NYSE 7/6 大反彈 +4.06%）

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

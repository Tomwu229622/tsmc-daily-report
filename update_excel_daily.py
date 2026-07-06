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

today_str = "2026-07-06"
tw_price = "NT$2,445（7/3收）"   # 台股2330 7/3 Fri官方收NT$2,445(-20、-0.81% vs 7/2收NT$2,465、量26,455張、TWSE確認)為最新確認收盤；跌幅較7/2之-1.60%收斂、守穩NT$2,410支撐；7/6 Mon盤中官方收盤待TWSE
change_pct = "-0.81%（7/3收）"    # 台股7/3收盤漲跌幅(-0.81%、跌幅收斂、守穩NT$2,410)
nyse_price = "US$434.16（7/2收）" # NYSE TSM 7/2 Thu官方收$434.16(-$10.07、-2.27%、stockanalysis.com確認、前收7/1 $444.23)為最新確認收盤——連二跌(7/1 -6.98%、7/2 -2.27%)但跌幅大幅收斂、殺盤趨緩；NYSE 7/3 Fri因美國國慶(7/4週六順延)休市、7/6 Mon重新開盤【重要更正】前一份7/6報告曾誤將7/2「更正」為+0.13%收$434.71，經官方收盤複核為誤、正確為-2.27%收$434.16
volume = "26,455（7/3收）"        # 台股7/3收盤成交量(張、TWSE確認、含一般+盤後定價)
news_summary = "報告日2026-07-06(Mon)。修正收斂、分析師大幅續挺：⚖️台股2330 7/3官方收NT$2,445(-20、-0.81%、量26,455張)跌幅較7/2之-1.60%收斂、守穩NT$2,410；NYSE TSM 7/2續跌-2.27%收$434.16(前收7/1 $444.23)為連二跌(7/1 -6.98%、7/2 -2.27%)但跌幅大幅收斂、殺盤動能衰竭、屬估值修正非基本面惡化；NYSE 7/3因美國國慶休市、7/6重新開盤。【重要更正】前一份7/6報告曾誤將7/2「更正」為+0.13%收$434.71，經stockanalysis.com官方收盤複核為誤、正確為-2.27%收$434.16。🚀Barclays將TSM目標價自$470大幅上調至$625(+43.9%空間)、Nomura NT$3,425(+40.1%空間)；平均目標~$490強力買入；⚠️惟Goldman 7/1移出APAC Conviction List(戰術調整非降評)。🤝NVIDIA-SK海力士宣布多年期記憶體合作(AI工廠次世代記憶體)；NVIDIA-TSMC導入AI入廠提升良率；BofA上修2026全球半導體市場至$1.3兆、生成式AI晶片估~$500B、費半YTD+47%。⭐漲價續發酵、7nm以下全先進製程調漲5-10%(佔晶圓營收~75%)、估全年毛利率+2pp以上。中長線基底未變：NVIDIA #1客戶(22%/$33B)/A16首發、Vera Rubin 6顆晶片全TSMC代工(3nm+HBM4)、2nm 70-80%良率。⚠️估值警示TSM P/E~37.8x、台股~36.0x；6月營收(7月上旬)與7/16財報為關鍵catalyst；ADR溢價約+10.3%；下方支撐NT$2,410/NT$2,370、上方壓力NT$2,465/NT$2,505/NT$2,535史高。"
change_color = "FF0000"        # 紅色（台股 7/3 收盤下跌 -0.81%；NYSE 7/2 續跌 -2.27% 但跌幅收斂、7/3 國慶休市、7/6 重新開盤）

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

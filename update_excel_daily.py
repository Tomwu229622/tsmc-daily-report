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

today_str = "2026-07-08"
tw_price = "NT$2,490（7/7收）"   # 台股2330最新確認收盤仍為7/7 Tue官方收NT$2,490(+30、+1.22% vs 7/6收NT$2,460、量估約24,800張、盤中高觸NT$2,505)；7/8 Wed盤中承壓補跌約NT$2,465(開2,465、區間2,455-2,500、約-1%)、跟進NYSE 7/7 -4.25%回檔、官方收盤待TWSE
change_pct = "+1.22%（7/7收）"    # 台股最新確認收盤漲跌幅為7/7 +1.22%(連二漲)；7/8盤中約-1%承壓補跌
nyse_price = "US$432.57（7/7收）" # NYSE TSM 7/7 Tue官方收$432.57(-$19.22、-4.25%、stockanalysis.com確認、前收7/6 $451.79、日內$428.11-$439.80)為最新確認NYSE收盤——回吐7/6假期後V型反彈(+4.06%)大部分漲幅、費半半導體同步回檔、7/16 Q2財報前估值修正+獲利了結、AI晶片股全面回落並非基本面惡化；NYSE 7/8盤待今晚GMT+8
volume = "24,800（7/7收，估）"    # 台股7/7收盤成交量(張、估)；7/8盤中量能待確認
news_summary = "報告日2026-07-08(Wed)。NYSE V型反彈後7/7回吐-4.25%、台股7/8恐承壓補跌：📉NYSE TSM 7/7收$432.57(-$19.22、-4.25%、前收7/6 $451.79、日內$428.11-$439.80、stockanalysis.com確認)回吐7/6假期後V型反彈(+4.06%)大部分漲幅、費半半導體同步回檔、為7/16 Q2財報前估值修正+獲利了結、AI晶片股全面回落並非基本面惡化；⚖️台股2330最新確認收盤仍為7/7 NT$2,490(+1.22%、連二漲、盤中高觸NT$2,505)、7/8盤中承壓補跌約NT$2,465(-1%)、料跟進NYSE回檔、關鍵守穩NT$2,460/NT$2,445/NT$2,410支撐；📊台美ADR溢價自7/6 +16.2%收斂至約+11.2%(7/7 ADR $432.57折合台股約NT$2,768 vs 台股7/7 NT$2,490)。🚀分析師仍大幅續挺:Citi NT$3,800(+32%、自NT$2,490具+52.6%空間)、Barclays維持$625(自$451.79具+38.3%空間)、Nomura NT$3,425(自NT$2,490具+37.6%空間)、共識強力買入;⚠️惟Goldman 7/1移出APAC Conviction List(戰術調整非降評)。🤝NVIDIA-SK海力士宣布多年期記憶體合作;NVIDIA-TSMC導入AI入廠提升良率;BofA上修2026全球半導體市場至$1.3兆(生成式AI晶片~$500B);Qualcomm推無HBM資料中心晶片。⭐漲價續發酵、7nm以下全先進製程調漲5-10%(佔晶圓營收~75%)、估全年毛利率+2pp以上。中長線基底未變:NVIDIA #1客戶(22%/$33B)/A16首發、Vera Rubin 6顆晶片全TSMC代工(3nm+HBM4)、2nm 70-80%良率、TSMC-Amkor美國封裝長約。⚠️估值警示TSM P/E~37.6x、台股~36.6x;4-5月營收+24% YoY落後~35%預期、6月營收(7月上旬)與7/16財報為關鍵catalyst;下方支撐NT$2,460/NT$2,445/NT$2,410、上方壓力NT$2,505/NT$2,535史高。"
change_color = "00B050"        # 綠色（台股最新確認收盤為7/7 +1.22%連二漲；惟7/8盤中承壓補跌、NYSE 7/7 -4.25%回檔）

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

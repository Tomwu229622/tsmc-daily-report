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

today_str = "2026-07-13"
tw_price = "NT$2,415.5（7/9收）"   # 台股2330最新確認收盤7/9 Thu官方收NT$2,415.5(-50、-2.03% vs 7/8 NT$2,465、開2,450/高2,460/低2,415、量27,057張、成交額658.50億、TWSE確認)失守5MA；7/10 Fri颱風「巴威」台股全日休市；7/13 Mon復市、6月營收今日13:30公布
change_pct = "-2.03%（7/9收）"    # 7/9官方漲跌幅-2.03%（TWSE、外資賣超12,749張拖累）；7/10颱風休市
nyse_price = "US$434.11（7/10收）" # NYSE TSM 7/10 Fri收$434.11(-$2.85、-0.65% vs 7/9 $436.96、stockanalysis.com確認、日內$428.10-$439.66)相對抗跌；7/9收$436.96(持平)；台股7/10休市期間NYSE續交易
volume = "27,057（7/9收）"    # 台股7/9官方成交量27,057張、成交額658.50億（TWSE）；7/10颱風休市無交易
news_summary = "報告日2026-07-13(Mon)。⭐⭐今日最重要：TSMC 6月營收今日13:30公布(原訂7/10、颱風順延)——法人估NT$4,000-4,400億、樂觀看逼近4,400億改寫單月新高，達標則Q2營收突破NT$1.2兆(對應指引$39.0-40.2B)；4-5月僅+24% YoY落後華爾街~35%預期、6月為Q2達標關鍵驗證、亦為7/16法說會前市場最大焦點。⚠️7/10 Fri中度颱風「巴威(Bavi)」侵台，證交所宣布台股(含期貨、外匯)全日休市；台股7/13復市。📉台股2330最新確認收盤7/9 Thu官方收NT$2,415.5(-50、-2.03%、開2,450/高2,460/低2,415、量27,057張、TWSE)失守5MA——主因外資單日大賣12,749張、提款約310億元、全集中市場連6日賣超(7/9賣471億)；投信買超43張(連7買)、自營買超90張、合計賣超12,615張。📊NYSE TSM台股休市期間續交易、相對抗跌：7/9收$436.96(持平)、7/10收$434.11(-0.65%)；NVDA 7/10收$210.96續強、AI晶片股維持多頭。📊台美ADR溢價升至+15.4%($434.11 vs理論$376.2、匯率~32.1)——溢價偏高部分因台股7/9起停格、NYSE續走。🚀分析師續挺：Citi NT$3,800(+57.3%)、Barclays $625(+44.0%)、Nomura NT$3,425(+41.8%)、共識強力買入；⚠️Goldman 7/1移出APAC Conviction List(戰術調整)；Motley Fool看好7/16財報後大漲。🎯催化倒數：6月營收今日13:30、7/16 Q2法說會倒數3天(Citi預期上修全年展望)。⭐中長線基底未變：NVIDIA #1客戶(22%/$33B)/A16首發、Vera Rubin 6顆晶片全TSMC代工、2nm 70-80%良率、TSMC-Amkor美國封裝長約、全先進製程漲價5-10%(毛利率+2pp)。⚠️風險：外資7月大幅賣超未轉向、6月營收若不如預期恐引賣壓、Intel 18A量產、下游手機/PC需求疲弱、估值不低(台股P/E~32.0x TTM、TSM~37x)、台幣~32.1。技術面：支撐NT$2,400/2,380/20MA~2,410、壓力NT$2,440/2,465/2,500-2,535史高。"
change_color = "FF0000"        # 紅色（7/9官方-2.03%回檔、外資大賣）

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

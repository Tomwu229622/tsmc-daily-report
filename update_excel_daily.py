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

today_str = "2026-07-01"
tw_price = "NT$2,455"         # 6/30 Tue 官方收盤(+85、+3.59% vs 6/29 NT$2,370、開2,440/高2,475/低2,435、量15,514張、成交額380.9億)；7/1台股盤中待TWSE
change_pct = "+3.59%"          # 6/30 官方收盤漲跌幅（跳空跟進NYSE前波大漲、突破5MA/20MA）
nyse_price = "US$468.80"       # NYSE TSM 6/30 Tue 收盤（最新確認、+3.02%、+$13.70、前收$455.10、逼近52週高$476.79、市值約$2.15兆、P/E~33.5）；7/1美股盤待今晚開盤
volume = "15,514"             # 6/30 官方全日成交量（張、成交額NT$380.9億、振幅1.69%）
news_summary = "報告日2026-07-01(Wed)；🚀今日頭條——台美雙市連袂走高、法人上調目標接力：🚀NYSE TSM 6/30 Tue續漲+3.02%(+$13.70)收$468.80、逼近52週高$476.79、市值約$2.15兆(P/E~33.5、前收$455.10)、7/1美股盤待今晚開盤;🚀台股2330 6/30官方收盤NT$2,455(+85、+3.59% vs 6/29 NT$2,370、開2,440/高2,475/低2,435、量15,514張、成交額380.9億、振幅1.69%)跳空跟進NYSE前波大漲、突破5MA/20MA、7/1台股盤中待TWSE。催化：📈Morgan Stanley上調目標並估2026營收+40% YoY;📈UBS上調目標並上修長期銷售成長;⭐漲價定價力擴及全先進製程(佔晶圓營收~74%、幅度5-10%、估全年毛利率+2pp以上);⭐TSMC-Amkor美國封裝長約降地緣風險。📊分析師共識目標升至平均$476.24、最高$625、最低$351。📊台美ADR溢價(以6/30 ADR$468.80+台股6/30 NT$2,455、匯率31.05計、理論值$395.33)走闊至+18.59%、台股7/1具補漲空間。🛠️晶片設備股領軍反彈：AMAT +10.8%、LRCX +8.4%、SOXX +4.1%(AI記憶體需求勁揚);Gartner估2026半導體營收超$1.3兆。⚠️漲價雙面刃——客戶(Nvidia/AMD/Apple/Qualcomm/Broadcom/MediaTek)成本升、Benzinga指NVIDIA因AI需求剛性+轉嫁能力強長期反為利多。⚠️近期營收疑慮續存——TSMC 4-5月營收合計+24% YoY低於華爾街~35%預期、6月營收(7月上旬)為關鍵catalyst;5月營收NT$4,169.75億(+30.1%)創單月新高、1-5月累計NT$1.96兆。⚠️ITC專利調查初判預計近期(6月底/7月初)出爐、為短線關鍵雜訊。⚠️監理雜訊——Super Micro遭台灣搜索(涉NVIDIA晶片走私中國)股價-8%;南韓公布逾$1兆美元AI/晶片投資、三星+SK 800兆韓元。⭐中長線基底未變：TSMC-Amkor 10年美國先進封裝長約(Peoria AZ$70億、2028投產)+Winbond特殊DRAM自主供應+NVIDIA確認第1大客戶(22%/$33B)/A16首發2027量產+NVIDIA-TSMC深化AI製程合作+2nm量產70-80%良率領先全球+2nm/A16產能2026-28估CAGR~70%+4大美國CSP 2026 AI CapEx合計$725B。⚠️風險續追蹤：(1)TSMC遭美ITC專利調查(Longitude/Marlin控訴)、近期預計初判、恐對含AI加速器技術晶片發布美國進口禁令;(2)Apple與Intel合作在美生產晶片、Intel 18A-P風險量產+UMC結盟、Google/AMD/Tesla接觸Samsung。基本面：Q1 2026淨利NT$572.5B(+58.3% YoY)、毛利率66.2%雙創史高、2025全年EPS NT$66.25。分析師全面看多、MS/UBS等上調目標(共識平均$476.24、最高$625)。⚠️估值警示TSM P/E~33.5x、台股P/E~36.2x；下方支撐NT$2,400/NT$2,370、上方壓力NT$2,490/NT$2,535史高。下一里程碑：ITC初判近期、6月營收7月上旬(關鍵)、Q2法說會7/16。"
change_color = "00B050"        # 綠色（6/30 官方收盤上漲 +3.59%）

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

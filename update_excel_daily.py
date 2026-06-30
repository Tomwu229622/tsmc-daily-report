"""
TSMC Excel 每日更新腳本 - 2026-06-30
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

today_str = "2026-06-30"
tw_price = "NT$2,370"          # 6/29 Mon 官方收盤（+1.28%、+30、量縮38,133張、成交額903.37億、開2,330/高2,395、自6/26 NT$2,340反彈）；6/30今日料受隔夜ADR激勵開高、官方收盤待TWSE
change_pct = "+1.28%"          # 6/29 漲跌幅（NYSE隔夜+5.26%帶動+Winbond DRAM合作+漲價定價力+BofA上調）
nyse_price = "US$455.10"       # NYSE TSM 6/29 Mon 收盤（+5.26%、+$22.75、前收$432.35、創波段新高、市值約$2.1兆、P/E~32.5）；Winbond DRAM合作+漲價定價力+BofA上調帶動
volume = "38,133"             # 6/29 官方成交量（張、量縮、成交額903.37億、振幅2.78%）
news_summary = "報告日2026-06-30(Tue)；🚀今日頭條——TSMC雙市同步反彈、漲價題材升級為定價力+自主供應鏈雙主軸：🚀NYSE TSM 6/29 Mon大漲+5.26%(+$22.75)收$455.10創波段新高、市值約$2.1兆(P/E~32.5、前收$432.35)；📈台股2330 6/29 Mon收NT$2,370(+30、+1.28%、量縮38,133張、成交額903.37億、開2,330/高2,395)自6/26 NT$2,340反彈；台股6/30今日料受隔夜ADR激勵開高、官方收盤待TWSE。三大催化：⭐TSMC與華邦電(Winbond)合作將CUBE特殊DRAM導入WoW 3D晶圓堆疊/SoIC先進封裝、把關鍵記憶體留台、降低對Samsung/SK海力士HBM依賴(在地化估縮短週期66%、降運費90%);⭐漲價定價力擴及先進+成熟製程(佔晶圓營收~75%、幅度5-10%、估全年毛利率+2pp以上);📈Bank of America等多家機構上調目標(彙整平均$468.08、最高$600、最低$351)。📊台美ADR溢價(以6/29 ADR$455.10+台股NT$2,370、匯率31.05計)走闊至+19.25%、ADR領先、台股具補漲空間。⚠️漲價雙面刃——客戶(Apple/NVIDIA/AMD/Qualcomm/Broadcom/MediaTek)成本升、Benzinga指NVIDIA因AI需求剛性+轉嫁能力強長期反為利多。⚠️近期營收疑慮續存——TSMC 4-5月營收合計+24% YoY低於華爾街~35%預期、6月營收(7月上旬)為關鍵catalyst;5月營收NT$4,169.75億(+30.1%)創單月新高、1-5月累計NT$1.96兆。⚠️ITC專利調查初判預計本週(6月底)出爐、為短線關鍵雜訊。⭐中長線基底未變：Winbond特殊DRAM自主供應+TSMC-Amkor 10年美國/韓國先進封裝長約(Peoria AZ$70億、2028投產)+NVIDIA確認第1大客戶(22%/$33B)/A16首發2027量產+NVIDIA-TSMC深化AI製程合作+2nm量產70-80%良率領先全球+2nm/A16產能2026-28估CAGR~70%+4大美國CSP 2026 AI CapEx合計$725B。⚠️風險續追蹤：(1)TSMC遭美ITC專利調查(Longitude/Marlin控訴)、6月底(本週)預計初判、恐對含AI加速器技術晶片發布美國進口禁令;(2)Apple與Intel合作在美生產晶片、Intel 18A-P風險量產+UMC結盟、Google/AMD/Tesla接觸Samsung。基本面：Q1 2026淨利NT$572.5B(+58.3% YoY)、毛利率66.2%雙創史高、2025全年EPS NT$66.25。分析師全面看多、BofA等上調目標(彙整平均$468.08、最高$600)。⚠️估值警示TSM P/E~32.5x、台股P/E~35.0x；下方支撐NT$2,340/NT$2,310、上方壓力5MA NT$2,396/NT$2,450/NT$2,535史高。下一里程碑：ITC初判6月底(本週)、6月營收7月上旬(關鍵)、Q2法說會7/16。"
change_color = "00B050"        # 綠色（6/29 上漲 +1.28%）

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

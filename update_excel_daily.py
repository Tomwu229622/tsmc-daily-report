"""
TSMC Excel 每日更新腳本 - 2026-06-10
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

today_str = "2026-06-29"
tw_price = "NT$2,340"          # 6/26 Fri 官方收盤（-2.09%、-50、量放大53,800張、外資賣超6,104張、自6/25 NT$2,390回落）；6/29今日開盤約NT$2,350小反彈、官方收盤待TWSE
change_pct = "-2.09%"          # 6/26 漲跌幅（隔夜漲價漲幅回吐+營收疑慮+外資賣超拖累）
nyse_price = "US$432.35"       # NYSE TSM 6/26 Fri 收盤（-0.61%、-$2.64、前收$434.99、市值約$2.0兆、P/E~31.0、日內$419.19–$436.13）；連續回吐6/24隔夜漲價漲幅
volume = "53,800"             # 6/26 官方成交量（張、量放大、賣壓增）
news_summary = "報告日2026-06-29(Mon)；⭐今日頭條——TSMC漲價題材仍為結構性主軸、惟上週五雙市回吐隔夜漲價漲幅：📉台股2330 6/26 Fri收NT$2,340(-50、-2.09%、量放大53,800張)自6/25 NT$2,390回落——隔夜ADR漲價漲幅回吐+近期營收疑慮+外資賣超(6/26外資賣超6,104張、投信買超179張、自營賣超310張、三大法人合計賣超6,235張)拖累。📉NYSE TSM 6/26 Fri收$432.35(-0.61%、-$2.64、前收$434.99、市值約$2.0兆、P/E~31.0、日內$419.19–$436.13)，連續回吐6/24漲價題材帶動之隔夜漲幅、估值與營收疑慮拉鋸。⭐漲價題材結構性利多未變：範圍確認擴大至7nm以下全部先進製程(含N7/N5/N3、佔晶圓營收約75%、N3單一佔25%)、幅度5-10%視客戶/節點/產品、部分已實施，以平均5%估可貢獻全年毛利率+2pp以上。📊台股6/29今日開盤約NT$2,350(+0.43% vs 6/26)小反彈、官方收盤待TWSE。📊台美ADR溢價(以6/26 ADR$432.35+台股NT$2,340、匯率31.05計)+14.74%。💡Benzinga：漲價雖壓縮NVIDIA毛利、惟AI需求剛性+轉嫁能力強、長期反為NVIDIA利多;⚠️價格敏感客戶(含Apple/手機SoC)影響較大、恐調整下單或評估替代。⚠️近期營收疑慮續存——TSMC 4-5月營收合計+24% YoY低於華爾街~35%預期、6月營收(7月上旬)為關鍵catalyst;5月營收NT$4,169.75億(+30.1%)創單月新高、1-5月累計NT$1.96兆。⚠️ITC專利調查初判預計本週(6月底)出爐、為短線關鍵雜訊。🛠️產業競局升溫：Intel 18A-P進入風險量產+UMC結盟、AMD瞄準$1,200億市場、Qualcomm推Dragonfly;Samsung/SK海力士送樣第7代HBM4E;Applied Materials 6/25推新一代設備加速DRAM與AI先進封裝。⭐中長線基底未變：TSMC-Amkor 10年美國/韓國先進封裝長約(Peoria AZ$70億、2028投產)+NVIDIA確認第1大客戶(22%/$33B)/A16首發2027量產+NVIDIA-TSMC深化AI製程合作+2nm量產70-80%良率領先全球+2nm/A16產能2026-28估CAGR~70%+4大美國CSP 2026 AI CapEx合計$725B。⚠️風險續追蹤：(1)TSMC遭美ITC專利調查(Longitude/Marlin控訴)、6月底(本週)預計初判、恐對含AI加速器技術晶片發布美國進口禁令;(2)Apple與Intel合作在美生產晶片、Google/AMD/Tesla接觸Samsung。基本面：Q1 2026淨利NT$572.5B(+58.3% YoY)、毛利率66.2%雙創史高、2025全年EPS NT$66.25。32位分析師全數看多、平均目標NT$2,530(上行+8.12%自NT$2,340、Susquehanna 6/23上調$575、待7/16法說會重估)。⚠️估值警示TSM P/E~31.0x、台股P/E~34.5x；下方支撐NT$2,310/NT$2,250、上方壓力NT$2,390/NT$2,450/NT$2,535史高。下一里程碑：ITC初判6月底(本週)、6月營收7月上旬(關鍵)、Q2法說會7/16。"
change_color = "FF0000"        # 紅色（6/26 下跌 -2.09%）

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

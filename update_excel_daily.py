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

today_str = "2026-06-26"
tw_price = "NT$2,390"          # 6/25 Thu 收盤（+15，+0.63%；跟進 6/24 ADR 反彈力道有限、OHLC/量為估算）；6/26 今日交易中、料跟進 ADR 盤後大漲、官方收盤待 TWSE
change_pct = "+0.63%"          # 6/25 漲跌幅（小幅反彈、漲價利多 vs 估值/營收疑慮拉鋸）
nyse_price = "US$440.83"       # NYSE TSM 6/25 Thu 收盤（+1.02%、+$4.44、漲價題材延燒、市值約$2.05兆、P/E~31.6）；盤後再+2.48%至$451.76、為最新確認數據
volume = "~14,000"            # 6/25 成交量（張，估）
news_summary = "報告日2026-06-26(Fri)；⭐今日頭條——TSMC晶圓代工漲價題材延燒、ADR續漲且盤後大漲、台股料補漲：⭐漲價範圍確認擴大至7nm以下全部先進製程(含N7/N5/N3、佔Q1晶圓營收約74%，N3單一佔25%)、幅度5-10%視客戶/節點/產品、部分已實施，以平均5%估可貢獻全年毛利率+2pp以上。📈NYSE TSM 6/25 Thu續漲+1.02%收$440.83(+$4.44、市值約$2.05兆、P/E~31.6)，盤後再大漲+2.48%至$451.76(漲價題材+產能擴張利多、今年規劃9期廠)。📊台股2330 6/25 Thu小漲+0.63%收NT$2,390、跟進6/24 ADR反彈力道有限；6/26今日台股料跟進ADR盤後大漲、官方收盤待TWSE。📊台美ADR溢價(以6/25 ADR$440.83+台股NT$2,390計)+14.54%、台股短線料補漲收斂。💡Benzinga：漲價雖壓縮NVIDIA毛利、惟AI需求剛性+轉嫁能力強、長期反為NVIDIA利多;⚠️價格敏感客戶(含Apple/手機SoC)影響較大、恐調整下單或評估替代。🛠️Applied Materials 6/25推新一代設備加速DRAM與AI先進封裝(HBM/3D堆疊)。⚠️近期營收疑慮續存——TSMC 4-5月營收合計+24% YoY低於華爾街~35%預期、6月營收(7月上旬)為關鍵catalyst;5月營收NT$4,169.75億(+30.1%)創單月新高、1-5月累計NT$1.96兆。⭐中長線基底未變：TSMC-Amkor 10年美國先進封裝長約(Peoria AZ$70億、2028投產)+NVIDIA確認第1大客戶(22%/$33B)/A16首發2027量產+NVIDIA-TSMC深化AI製程合作+2nm量產70-80%良率領先全球+4大美國CSP 2026 AI CapEx合計$725B+2026全球半導體市場估+26% YoY達~$975B。⚠️風險續追蹤：(1)TSMC遭美ITC專利調查(Longitude/Marlin控訴)、6月底預計初判、恐對含AI加速器技術晶片發布美國進口禁令;(2)Apple與Intel合作在美生產晶片、Google/AMD/Tesla接觸Samsung、Intel 18A-P風險量產+UMC結盟。基本面：Q1 2026淨利NT$572.5B(+58.3% YoY)、毛利率66.2%雙創史高、2025全年EPS NT$66.25。32位分析師全數看多、平均目標NT$2,530(上行+5.86%自NT$2,390、Susquehanna 6/23上調$575、待7/16法說會重估)。⚠️估值警示TSM P/E~31.6x、台股P/E~35.2x；下方支撐NT$2,350/NT$2,310、上方壓力NT$2,450/NT$2,490/NT$2,535史高。下一里程碑：ITC初判6月底、6月營收7月上旬(關鍵)、Q2法說會7/16。"
change_color = "00B050"        # 綠色（6/25 上漲）

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

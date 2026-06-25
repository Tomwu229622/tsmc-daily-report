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

today_str = "2026-06-25"
tw_price = "NT$2,390"          # 6/24 Wed 官方收盤（-100，-4.02% vs 6/23 NT$2,490、開 2,435 高 2,445 低 2,390 收最低、量 12,657 張、市值 NT$62.0 兆、補跌跟進 6/23 ADR 暴跌）；6/25 今日交易中、料跟進 ADR 反彈、官方收盤待 TWSE
change_pct = "-4.02%"          # 6/24 漲跌幅（補跌跟進 6/23 ADR 暴跌、半導體全面拋售）
nyse_price = "US$441.35"       # NYSE TSM 6/24 收盤（+1.14%、+$4.96 vs 6/23 $436.39、受晶圓代工漲價 5-10% 利多反彈、日內 $432.58-$443.85、市值約$2.05兆、P/E~31.6、為最新確認數據）
volume = "12,657"             # 6/24 成交量（張）
news_summary = "報告日2026-06-25(Thu)；⭐今日頭條——TSMC全面調漲晶圓代工報價5-10%、ADR反彈但台股補跌：⭐TSMC對7nm以下所有先進製程(佔營收約75%)全面調漲5-10%報價、以平均5%估算可貢獻全年毛利率+2pp以上，消息帶動NYSE TSM 6/24 Wed反彈+1.14%收$441.35(+$4.96 vs 6/23 $436.39、日內$432.58-$443.85、市值約$2.05兆、P/E~31.6)。📉台股2330 6/24 Wed補跌-4.02%官方收NT$2,390(-100 vs 6/23 NT$2,490、開2,435高2,445低2,390收最低、量12,657張、市值NT$62.0兆)——跟進6/23 ADR暴跌-6.69%跳空走低、台股加權同日重挫-1,057點寫史上第8慘、半導體類股全面回檔；6/25今日台股交易中、料跟進ADR反彈、官方收盤待TWSE。📊台美ADR溢價(以6/24 ADR$441.35+台股NT$2,390計)自+8.83%回升至+14.68%、台股短線料補漲收斂。⚠️漲價雙面刃：利多毛利率，但傳客戶(含Apple)措手不及、恐促部分客戶調整下單量或轉向UMC等成熟製程替代。⚠️近期營收疑慮續存——TSMC 4-5月營收合計+24% YoY低於華爾街~35%預期、6月營收(7月上旬)為關鍵catalyst。📈Susquehanna將TSM目標價自$500上調至$575。⭐中長線基底未變：TSMC-Amkor簽10年美國先進封裝長約(6/16公告、Peoria AZ$70億、2028投產)+NVIDIA確認第1大客戶(22%/$33B)/A16首發2027量產+NVIDIA-TSMC深化AI製程合作(cuLitho/缺陷檢測/工廠排程)+RTX Spark超晶片由TSMC代工+2nm量產70-80%良率領先全球+CoPoS/CoWoS擴產+4大美國CSP 2026 AI CapEx合計$725B+2026全球半導體市場估+25% YoY達~$975B。⚠️風險續追蹤：(1)Apple同意與Intel合作在美生產晶片降低對TSMC依賴、Google/AMD/Tesla接觸Samsung、Intel 18A-P風險量產;(2)TSMC遭美ITC專利調查(Longitude/Marlin控訴)、6月底預計初判、恐對含AI加速器技術晶片發布美國進口禁令。基本面：Q1 2026淨利NT$572.5B(+58.3% YoY)、毛利率66.2%雙創史高、2025全年EPS NT$66.25。32位分析師全數看多、平均目標NT$2,530(上行+5.86%自NT$2,390、待7/16法說會重估)。⚠️估值警示TSM P/E~31.6x、台股P/E~35.2x、留意6月營收miss風險與台幣升值；下方支撐NT$2,350/NT$2,310、上方壓力NT$2,450/NT$2,490/NT$2,535史高。下一里程碑：6月營收7月上旬(關鍵)、Q2法說會7/16。"
change_color = "FF0000"        # 紅色（6/24 下跌）

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

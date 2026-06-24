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

today_str = "2026-06-24"
tw_price = "NT$2,490"          # 6/23 Tue 官方收盤（-20，-0.80% vs 6/22 NT$2,510、盤中觸史高 NT$2,535 後拉回、量 34,148 張、相對抗跌）；6/24 今日交易中、料跟進 ADR 暴跌跳空走低、官方收盤待 TWSE
change_pct = "-0.80%"          # 6/23 漲跌幅（自史高小幅獲利了結、相對抗跌）
nyse_price = "US$436.39"       # NYSE TSM 6/23 收盤（-6.69%、-$31.28 vs 6/22 $467.67、半導體全面拋售暴跌、市值約$1.97兆、P/E~31.3、為最新確認數據）
volume = "34,148"             # 6/23 成交量（張，官方）
news_summary = "報告日2026-06-24(Wed)；⚠️半導體全面拋售、ADR暴跌但台股抗跌：NYSE TSM 6/23暴跌-6.69%收$436.39(-$31.28 vs 6/22 $467.67、市值回落約$1.97兆、P/E~31.3)，半導體類股全面拋售——Micron -11.4%、AMD -6.25%、AVGO -5.02%、設備股跌4.5-5%；崩跌觸發：Broadcom指引不如預期引爆sell-the-news+BoFA升息報告推升美債殖利率+亞股下挫+AI過熱疑慮+財報前獲利了結，單日逾千億美元半導體市值蒸發。台股2330 6/23官方收NT$2,490(-20、-0.80% vs 6/22 NT$2,510、量34,148張、市值NT$64.6兆)相對抗跌、盤中觸歷史新高NT$2,535後拉回收最低；6/24今日台股交易中、料跟進ADR暴跌跳空走低、官方收盤待TWSE。📊台美ADR溢價(以6/23 ADR$436.39+台股NT$2,490計)自+15.71%大幅收斂至+8.83%、台股短線恐有補跌壓力。⚠️TSMC 4-5月營收合計+24% YoY低於華爾街~35%預期、近期營收恐miss、6月營收(7月上旬)為關鍵catalyst。📈逆勢利多：Susquehanna將TSM目標價自$500上調至$575。⭐中長線基底未變：TSMC-Amkor簽10年美國先進封裝長約(6/16公告、Peoria AZ$70億、2028投產)+NVIDIA確認第1大客戶(22%/$33B)/A16首發2027量產+NVIDIA-TSMC深化AI製程合作(cuLitho/缺陷檢測/工廠排程)+RTX Spark超晶片由TSMC代工+2nm量產70-80%良率領先全球+CoPoS/CoWoS擴產+4大美國CSP 2026 AI CapEx合計$725B+2026全球半導體市場估+25% YoY達~$975B。⚠️風險續追蹤：(1)Apple同意與Intel合作在美生產晶片降低對TSMC依賴、Google/AMD/Tesla接觸Samsung、Intel 18A-P風險量產;(2)TSMC遭美ITC專利調查(Longitude/Marlin控訴)、6月底預計初判、恐對含AI加速器技術晶片發布美國進口禁令。基本面：Q1 2026淨利NT$572.5B(+58.3% YoY)、毛利率66.2%雙創史高、2025全年EPS NT$66.25。32位分析師全數看多、平均目標NT$2,530(上行+1.61%、待7/16法說會重估)。⚠️估值警示TSM P/E~31.3x、台股P/E~36.7x、仍在5年90%+高位、留意6月營收miss風險與台幣升值；上方壓力NT$2,510/NT$2,535史高、下方支撐NT$2,450/NT$2,410。下一里程碑：6月營收7月上旬(關鍵)、Q2法說會7/16。"
change_color = "FF0000"        # 紅色（6/23 下跌）

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

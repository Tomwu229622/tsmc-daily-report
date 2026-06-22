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

today_str = "2026-06-22"
tw_price = "NT$2,410"          # 6/18 Thu 官方收盤（+25，+1.05% vs 6/17 NT$2,385、盤中觸 NT$2,440 史高、量 49,982 張）；6/19 端午休市、6/22 今日交易中、官方收盤待 TWSE
change_pct = "+1.05%"          # 6/18 漲跌幅（台股加權指數同日 +901.85 點收 46,459 連 4 日創新高）
nyse_price = "US$462.12"       # NYSE TSM 6/18 收盤（+6.94%、+$29.97 vs $432.15、盤中觸史高 $465.22、為最新確認數據）
volume = "49,982"             # 6/18 成交量（張，官方）
news_summary = "報告日2026-06-22(Mon)；🚀NYSE TSM 6/18創紀錄飆漲收$462.12(+6.94%、+$29.97 vs $432.15)、盤中觸52週/歷史新高$465.22、市值$1.98兆、量2,582萬股、P/E 33.14(最新確認)；台股2330 6/18官方收NT$2,410(+25、+1.05% vs 6/17 NT$2,385、盤中觸NT$2,440史高、量49,982張)、加權指數+901.85點收46,459連4日創新高(前報估值2,385已更正)；6/19端午休市、6/22今日交易中受ADR飆漲帶動預期跳空跟漲、官方收盤待TWSE。🤝本波飆漲主因TSMC-Amkor簽10年美國先進封裝長約(6/16公告、Peoria AZ廠$20億→$70億、2028投產、~2,000就業、InFO/CoWoS、呼應$165B美國投資)+分析師升評+AI需求。🏆NVIDIA確認TSMC第1大客戶(22%/$33B)、為A16(1.6nm)首發客戶2027量產；Apple跳過A16直上A14。⚠️新增風險：(1)Apple同意與Intel合作在美生產晶片降低對TSMC依賴、Google/AMD/Tesla接觸Samsung；Intel飆漲+10.6%創新高；(2)TSMC遭美ITC專利調查(Longitude/Marlin控訴)、6月預計初判、恐對含AI加速器技術晶片發布美國進口禁令。📊台美ADR溢價(以6/18 ADR$462.12+台股NT$2,410計)大幅擴大至+19.08%、6/22台股有補漲壓力。⭐中長線基底未變：5月營收NT$4,169.75億+30.1%創單月新高、2nm量產70-80%良率領先全球、4大美國CSP 2026 AI CapEx合計$725B、SIA/Deloitte估AI資料中心晶片營收2028上看$1.2兆(4年近10倍)、2026全球半導體市場估+25% YoY達~$975B。基本面：Q1 2026淨利NT$572.5B(+58.3% YoY)、毛利率66.2%雙創史高、2025全年EPS NT$66.25。32位分析師全數看多、平均目標NT$2,530(上行+4.98%)。⚠️估值警示TSM P/E 33.1x、台股P/E~35.5x、留意台幣升值與短線過熱；上方壓力NT$2,440史高/NT$2,500、下方支撐NT$2,410/NT$2,375。下一里程碑：6月營收7月上旬、Q2法說會7/16。"
change_color = "00B050"        # 綠色（6/18 上漲）

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

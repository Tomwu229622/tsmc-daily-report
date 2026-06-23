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

today_str = "2026-06-23"
tw_price = "NT$2,510"          # 6/22 Mon 官方收盤（+100，+4.15% vs 6/18 NT$2,410、收最高、創歷史新高、量 39,134 張）；6/19 端午休市、6/23 今日交易中、開盤 NT$2,455、官方收盤待 TWSE
change_pct = "+4.15%"          # 6/22 漲跌幅（受 6/18-6/19 NYSE 飆漲帶動跳空大漲、創史高）
nyse_price = "US$467.67"       # NYSE TSM 6/22 收盤（+1.20%、+$5.55 vs 6/19 ~$462.12、再創收盤紀錄新高、市值約$2.4兆、為最新確認數據）
volume = "39,134"             # 6/22 成交量（張，官方）
news_summary = "報告日2026-06-23(Tue)；🚀台美雙市再創歷史新高：台股2330 6/22官方收NT$2,510(+100、+4.15% vs 6/18 NT$2,410、收最高、量39,134張、市值NT$65.1兆)創史高，受6/18 NYSE +6.94%飆漲與6/19美股交易帶動跳空大漲；NYSE TSM 6/22收$467.67(+$5.55、+1.20% vs 6/19~$462.12)再創收盤紀錄新高、市值約$2.4兆、P/E~33.5(最新確認)；6/19端午休市、6/23今日台股開盤NT$2,455自史高小幅拉回、交易中、官方收盤待TWSE。📊台美ADR溢價(以6/22 ADR$467.67+台股NT$2,510計)收斂至+15.71%(自6/18 +19.08%)、台股已補漲跟上。🤝本波飆漲主因TSMC-Amkor簽10年美國先進封裝長約(6/16公告、Peoria AZ廠$20億→$70億、2028投產、~2,000就業、InFO/CoWoS、呼應$165B美國投資)+分析師升評+AI需求。📦新動態：TSMC加速CoPoS面板級封裝布局、目標6月底完成材料/設備驗證、擴大AI封裝產能；惟Samsung PLP技術仍領先數年。🏆NVIDIA確認TSMC第1大客戶(22%/$33B)、為A16(1.6nm)首發客戶2027量產、6月完成$250億債券發行；Apple跳過A16直上A14。⚠️風險續追蹤：(1)Apple同意與Intel合作在美生產晶片降低對TSMC依賴、Google/AMD/Tesla接觸Samsung、Intel與UMC結盟；(2)TSMC遭美ITC專利調查(Longitude/Marlin控訴)、6月預計初判、恐對含AI加速器技術晶片發布美國進口禁令。⭐中長線基底未變：5月營收NT$4,169.75億+30.1%創單月新高、2nm量產70-80%良率領先全球、4大美國CSP 2026 AI CapEx合計$725B、SIA/Deloitte估AI資料中心晶片營收2028上看$1.2兆(4年近10倍)、2026全球半導體市場估+25% YoY達~$975B。基本面：Q1 2026淨利NT$572.5B(+58.3% YoY)、毛利率66.2%雙創史高、2025全年EPS NT$66.25。32位分析師全數看多、平均目標NT$2,530(已趨近、上行+0.80%、待法說會上修)。⚠️估值警示TSM P/E~33.5x、台股P/E~37.0x、留意台幣升值與短線過熱(RSI 74/KDJ J96進入超買)；上方壓力NT$2,510史高/NT$2,550、下方支撐NT$2,455/NT$2,410。下一里程碑：6月營收7月上旬、Q2法說會7/16。"
change_color = "00B050"        # 綠色（6/22 上漲）

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

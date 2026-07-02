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

today_str = "2026-07-02"
tw_price = "NT$2,505"         # 7/1 Wed 官方收盤(+95、+3.94% vs 6/30 NT$2,410、高2,505、量37,544張)連3漲、下半年開門紅、突破5MA/20MA；7/2台股盤中待TWSE
change_pct = "+3.94%"          # 7/1 官方收盤漲跌幅（連3漲、下半年開門紅、突破5MA/20MA）
nyse_price = "US$444.23"       # NYSE TSM 7/1 Wed 收盤（最新確認、-6.98%、-$33.34、前收$477.57、自6/30史新高$477.57重挫、盤後回升+0.94%至$448.41、市值約$2.30兆、P/E~38.6）；7/2美股盤待今晚開盤
volume = "37,544"             # 7/1 官方全日成交量（張、量能明顯放大確認突破）
news_summary = "報告日2026-07-02(Thu)；⚠️今日頭條——台美劇烈背離、NYSE夜盤重挫：📉NYSE TSM 7/1 Wed重挫-6.98%(-$33.34)收$444.23、自6/30史新高$477.57大幅回落、市值約$2.30兆(P/E~38.6)、盤後回升+0.94%至$448.41顯示殺盤趨緩、7/2美股盤待今晚開盤;🚀台股2330 7/1官方收盤NT$2,505(+95、+3.94% vs 6/30 NT$2,410、高2,505、量37,544張)連3漲、下半年開門紅、突破5MA/20MA——惟台股7/1收盤在NYSE夜盤重挫之前、7/2恐面臨補跌壓力、盤中待TWSE。NYSE重挫催化：技術面過熱獲利了結(費半SOX Q2 +87.8%創1994年以來最強單季後拉回)、估值修正、4-5月營收+24% YoY落後華爾街~35%預期、智慧型手機/PC下游需求疲弱、7/16財報前觀望。📊台美ADR溢價(以7/1 ADR$444.23+台股7/1 NT$2,505、匯率31.05計、理論值$403.38)自+18.59%大幅收斂至+10.13%、預示台股7/2補跌。📊分析師共識仍偏多——17位Strong Buy、0賣出、12個月平均目標$487.56(自7/1收$444.23具+9.8%上行空間)、拉回視為估值修正非基本面惡化。⭐漲價題材續發酵——TSMC通知客戶全先進製程調漲5-10%(擴及N7/N5/N3及成熟製程、佔晶圓營收~74%)、估全年毛利率+2pp以上;客戶(Nvidia/AMD/Apple/Qualcomm/Broadcom/MediaTek)成本升、Benzinga指NVIDIA因AI需求剛性+轉嫁能力強長期反為利多。🏆NVIDIA確認超越Apple成TSMC最大客戶(22%/$33B)、A16首發2026下半/2027量產;Apple跳過A16直上A14。⚠️南韓公布逾$1兆美元AI/晶片投資、三星+SK 800兆韓元(約$5,180億)各建兩座新廠、中長期強化韓系競爭。⭐中長線基底未變：TSMC-Amkor 10年美國先進封裝長約(Peoria AZ$70億、2028投產)+Winbond特殊DRAM自主供應+NVIDIA第1大客戶/A16首發+2nm量產70-80%良率領先全球+2nm/A16產能2026-28估CAGR~70%+4大美國CSP 2026 AI CapEx合計$725B;全球半導體2026估$975B(+26% YoY)、AI資料中心晶片2028估$1.2兆(佔AI伺服器機櫃價值95%)。基本面：Q1 2026淨利NT$572.5B(+58.3% YoY)、毛利率66.2%雙創史高、2025全年EPS NT$66.25。⚠️估值警示TSM P/E~38.6x、台股P/E~36.9x；下方支撐NT$2,450/NT$2,410、上方壓力NT$2,505/NT$2,535史高。下一里程碑：6月營收7月上旬(關鍵catalyst)、Q2法說會7/16。"
change_color = "00B050"        # 綠色（台股 7/1 官方收盤上漲 +3.94%）

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

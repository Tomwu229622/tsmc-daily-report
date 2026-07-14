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

today_str = "2026-07-14"
tw_price = "NT$2,440（7/13收）"   # 台股2330最新確認收盤7/13 Mon官方收NT$2,440(+25、+1.04% vs 7/9 NT$2,415.5、開2,460/高2,480/低2,440、量34,665張、成交額851.13億、均價2,455.21、TWSE)站回5MA；颱風休市後首個交易日、6月營收創高撐盤、盤中一度大漲逾900點尾盤翻黑收僅+25
change_pct = "+1.04%（7/13收）"    # 7/13官方漲跌幅+1.04%（TWSE、6月營收創高達標帶動）
nyse_price = "US$421.58（7/13收）" # NYSE TSM 7/13 Mon收$421.58(-$12.53、-2.89% vs 7/10 $434.11、stockanalysis.com確認、自我檢核434.11×(1-2.89%)=421.56✓、盤後$422.98)法說會前獲利了結、半導體回檔
volume = "34,665（7/13收）"    # 台股7/13官方成交量34,665張、成交額851.13億（TWSE）
news_summary = "報告日2026-07-14(Tue)。⭐⭐今日最重要：TSMC 6月營收7/13正式公布大幅達標創高——6月合併營收NT$4,426.80億(+67.9% YoY、+6.2% MoM)改寫單月歷史新高；1-6月(H1)累計NT$2.40兆(+35.6% YoY)；Q2累計約NT$1.27兆、達成美元指引$39.0-40.2B——AI/HPC需求動能獲確認、大幅緩解4-5月僅+24% YoY落後~35%預期的營收疑慮、為7/16法說會前重大利多。📈台股2330 7/13 Mon官方收NT$2,440(+25、+1.04% vs 7/9 NT$2,415.5、開2,460/高2,480/低2,440、量34,665張、額851.13億、TWSE)站回5MA——颱風休市後首個交易日、6月營收創高撐盤；⚡盤中一度大漲逾900點、尾盤獲利了結翻黑收僅+25、台積電為撐盤要角、顯示高檔換手。📉NYSE TSM 7/13收$421.58(-2.89% vs 7/10 $434.11、盤後$422.98)——法說會前美股半導體獲利了結、估值修正。📊三大法人7/13：外資賣超2,045張(連3賣、3日累計-16,464張)、投信買超530張、自營買超623張、合計賣超892張——外資賣壓大幅收斂、投信自營同步買超對沖。📊台美ADR溢價自7/10約+15.3%收斂至約+10.8%($421.58×32.07÷5=NT$2,704 vs台股2,440)。🚀分析師續挺：Citi NT$3,800(+55.7%)、Barclays $625(+48.3%)、Nomura NT$3,425(+40.4%)、共識強力買入；⚠️Goldman 7/1移出APAC Conviction List(戰術調整)。🎯催化倒數：7/16 Q2法說會週三14:00倒數2天(公布Q2財報+Q3展望、法人估Q2毛利率挑戰70%、Citi預期上修全年展望)。⭐中長線基底未變：NVIDIA #1客戶(22%/$33B)/A16首發、Vera Rubin 6顆晶片全TSMC代工、2nm 70-80%良率、TSMC-Amkor美國封裝長約、全先進製程漲價5-10%(毛利率+2pp)。⚠️風險：外資7月連3賣未轉向、法說會前NYSE續弱、Intel 18A量產、下游手機/PC需求疲弱、估值不低(台股P/E~32.3x TTM、TSM~36x)、台幣~32.1。技術面：站回5MA~2,424、支撐20MA~2,366/2,300、壓力2,480/2,500/2,535史高。"
change_color = "00B050"        # 綠色（7/13官方+1.04%、6月營收創高達標帶動）

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

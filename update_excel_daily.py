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

today_str = "2026-07-09"
tw_price = "NT$2,465（7/8收）"   # 台股2330 7/8 Wed官方收NT$2,465(+25、+1.02% vs 7/7 NT$2,440、開2,445/高2,465/低2,420、量25,519張、TWSE確認、收在當日最高)；【更正】7/7官方收NT$2,440(-20、-0.81%)非前報之NT$2,490；7/9開盤NT$2,450盤中待收
change_pct = "+1.02%（7/8收）"    # 7/8官方漲跌幅+1.02%（TWSE）；台美7/8同步反彈同幅+1.02%
nyse_price = "US$436.98（7/8收）" # NYSE TSM 7/8 Wed收$436.98(+$4.41、+1.02% vs 7/7 $432.57、Yahoo Finance日線確認)自7/7 -4.25%回檔止穩；SOX +2.23%收12,574.97、NVDA +3.65%、AVGO +4.83% AI晶片股回神
volume = "25,519（7/8收）"    # 台股7/8官方成交量25,519張（25,519,599股、TWSE）
news_summary = "報告日2026-07-09(Thu)。✅台美7/8同步止跌反彈：台股2330官方收NT$2,465(+25、+1.02%、量25,519張、收在當日最高、站回5MA 2,455)；NYSE TSM收$436.98(+1.02%)自7/7 -4.25%回檔止穩、SOX +2.23%、NVDA +3.65%、AVGO +4.83% AI晶片股回神。⚠️【數據更正】經TWSE官方核實：7/7實收NT$2,440(-0.81%、盤中高2,500)非前報誤載之NT$2,490(+1.22%)；6/30官方收NT$2,410；7/7外資實為+83張(前報估+12,000有誤)——技術指標已重算(5MA 2,455/20MA 2,403)。📊三大法人7/8官方(T86)：外資-4,155張、投信+730、自營+468、合計-2,957張；7月外資累計約-2.9萬張、投信連5日買超。📊 ADR溢價+13.7%($436.98 vs理論$384.32、匯率32.07)。🚀分析師續挺：Citi NT$3,800(+54.2%)、Barclays $625(+43.0%)、Nomura NT$3,425(+38.9%)、共識強力買入；⚠️Goldman 7/1移出APAC Conviction List(戰術調整)。🎯催化倒數：6月營收7月上旬隨時公布(4-5月+24% YoY落後~35%預期、攸關Q2指引$39.0-40.2B達成)、7/16 Q2法說會倒數7天(Citi預期上修全年展望)。⭐中長線基底未變：NVIDIA #1客戶(22%/$33B)/A16首發、Vera Rubin 6顆晶片全TSMC代工、2nm 70-80%良率、TSMC-Amkor美國封裝長約、全先進製程漲價5-10%(毛利率+2pp)。⚠️風險：外資7月賣超未轉向、Intel 18A量產(微軟/高通進駐)、下游手機/PC需求疲弱、估值不低(台股P/E~32.7x TTM、TSM~37x)、台幣32.07。技術面：支撐NT$2,440/2,420/2,403、壓力NT$2,500-2,505/2,535史高。"
change_color = "00B050"        # 綠色（7/8官方+1.02%止跌回升、收在當日最高）

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

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

today_str = "2026-07-20"
tw_price = "NT$2,290（7/17收）"   # 台股2330 7/17 Fri官方收NT$2,290(-180、-7.29% vs 7/16 NT$2,470、開2,375/高2,395/低2,290、量97,362張TWSE口徑、成交額2,290.52億、成交1,150,086筆、TWSE確認)——史詩級崩跌、收盤即當日最低、AI晶片股拋售擴散至台股
change_pct = "-7.29%（7/17收）"    # 7/17官方漲跌幅-7.29%（TWSE、AI晶片股熊市拋售、外資史詩級提款）
nyse_price = "US$398.37（7/17收）" # NYSE TSM 7/17 Fri收$398.37(-$11.37、-2.77% vs 7/16 $409.74、盤後$397.72、stockanalysis.com確認)——失守$400；SOX正式跌入熊市(自6/22高-20.2%)、惟相對板塊抗跌
volume = "97,362（7/17收）"    # 台股7/17官方成交量97,362張（TWSE口徑、前日3.2倍）、成交額2,290.52億
news_summary = "報告日2026-07-20(Mon)。⚠️⚠️全球AI晶片股正式跌入熊市：SOX 7/17收-1.6%(盤中一度-5.7%)、自6/22史高回落-20.2%、單週-9%(3-6月曾自低點飆+105%)；催化：中國Moonshot發表號稱全球最大開源AI模型Kimi K3、Alphabet傳Gemini 3.5 Pro交付落後——AI投資報酬疑慮引爆全面獲利了結；MRVL/ARM/INTC自高點已-30%+。📉台股同步史詩級崩跌：2330 7/17官方收NT$2,290(-180、-7.29%、收盤即當日最低、量爆增97,362張、額2,290億、前日3.2倍)；三大法人7/17合計賣超逾2,631.5億元創台股史上單日最大、外資提款台積電破千億——2330遭外資賣超44,184張(連7賣、買17,479/賣61,663)；惟投信+1,489(連2買)、自營+2,759(連7買)逆勢承接、合計賣超39,936張。📊台美ADR溢價逆勢擴大至+12.0%($398.37 vs理論值$355.6、匯率~32.2)——台股-7.29%超跌於ADR-2.77%、隱含7/20週一具技術性反彈空間。✅關鍵：本波為AI估值/情緒修正、非基本面惡化——Q2法說會(7/16)大超預期完全未變：營收$40.20B(+33.7% YoY)、EPS$4.31/ADR(+77.4%、連11季超預期)、毛利率67.7%；全年成長上修>40%、CapEx上調$60-64B(含Arizona追加$100B)、股利+33%；管理層:「AI需求比預期更強」——股價與基本面短線脫鉤。💡估值大幅消化：台股P/E降至~26.2x TTM、TSM~29.2x。技術面：跌破5MA~2,412/20MA~2,428、回測60MA~2,324、RSI(14)41.0、MACD綠柱擴大至-15.2(DIF 17.4<DEA 32.5)、KDJ K35.7/D46.2/J14.7死叉墜超賣；支撐2,290/2,200/2,094、壓力2,322-2,324(布林下軌/60MA)/2,412/2,470。⚠️風險：AI情緒面能否止穩、外資連7史詩級賣超能否止穩為兩大變數；美232關稅25%+輸中晶片轉運查驗、Intel 18A競爭。"
change_color = "FF0000"        # 紅色（7/17官方-7.29%下跌）

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

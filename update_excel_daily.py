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

today_str = "2026-07-17"
tw_price = "NT$2,470（7/16收）"   # 台股2330 7/16 Thu官方收NT$2,470(+30、+1.23% vs 7/15 NT$2,440、開2,430/高2,470/低2,420、量30,539張TWSE口徑、成交額747.50億、均價2,448、TWSE)——大盤-0.01%下逆勢領漲、法說會前搶跑
change_pct = "+1.23%（7/16收）"    # 7/16官方漲跌幅+1.23%（TWSE、大盤-0.01%下逆勢領漲）
nyse_price = "US$409.74（7/16收）" # NYSE TSM 7/16 Thu收$409.74(-$9.74、-2.32% vs 7/15 $419.48、日內403.50-415.95、量2,479萬股、Yahoo)——Q2大超預期仍收跌：SOX -4.29%重挫、AI估值疑慮+H2毛利率稀釋+CapEx上調
volume = "30,539（7/16收）"    # 台股7/16官方成交量30,539張（TWSE口徑）、成交額747.50億
news_summary = "報告日2026-07-17(Fri)。⭐⭐Q2法說會(7/16 14:00)大超預期：營收NT$1.270兆($40.20B、+33.7% YoY、+12.0% QoQ)觸及指引上緣；EPS NT$27.25($4.31/ADR、+77.4% YoY)超共識~$3.94約+9.4%、連11季超預期、連5季獲利創高；毛利率67.7%超指引上緣、營益率60.3%、淨利率55.6%；Q3指引$44.6-45.8B(中位+12.4% QoQ、遠超預期)、毛利率65-67%(2nm爬坡H2稀釋3-4pp)；全年營收成長自>30%上修至「略高於40%」；CapEx自$52-56B上調至$60-64B；2026股利NT$24/股(+33%)；HPC占66%、2nm首度貢獻3%；管理層:「AI需求比先前預期更強」。⚠️惟美股不買單：NYSE TSM 7/16收$409.74(-2.32%)——SOX -4.29%收11,867.50全面重挫(AVGO -5.03%、AMD -5.33%、INTC -5.84%、MU -5.65%、NVDA -2.40%)，分析師點名重資本支出、H2毛利率稀釋、AI估值疑慮+美伊地緣緊張；記憶體鏈續崩(SK海力士-11.5%、三星-8.8%)；⭐AAPL +1.76%逆勢。📈台股2330 7/16收NT$2,470(+30、+1.23%、量30,539張、額747.50億)法說會前搶跑、大盤-0.01%下逆勢領漲——今日7/17開盤補反映「財報大利多vs美股重挫」多空對決、恐高波動。📊三大法人7/16(T86)合計賣超2,625張：外資-4,215(連6賣、近5交易日-34,852)、投信+528(轉買)、自營+1,061——內資先行卡位、外資能否翻買為最大觀察點。📊ADR溢價收斂至+7.0%($409.74 vs理論值$383.08、匯率32.24)。💡估值大幅消化：P/E降至~28.3x TTM(前~32.3x)、TSM ~30.0x。技術面：站上5MA~2,437/20MA~2,434、RSI 57.8、MACD柱收斂至-4.9、KDJ初步金叉；支撐2,420/2,415/2,349、壓力2,480/2,520/2,535史高。"
change_color = "00B050"        # 綠色（7/16官方+1.23%上漲）

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

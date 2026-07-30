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

today_str = "2026-07-30"
tw_price = "NT$2,200（7/29收）"   # 台股2330 7/29 Wed官方收NT$2,200(-80、-3.51% vs 7/28 NT$2,280、開2,260/高2,280/低2,180、量68,140張TWSE口徑、成交額1,515.52億、成交675,528筆、TWSE確認)——股災第2日、13:00後恐慌殺盤觸2,180尾盤拉回；跌幅與大盤-3.76%相當、抗跌角色弱化
change_pct = "-3.51%（7/29收）"    # 7/29官方漲跌幅-3.51%（TWSE、兩日累計-6.4%、連2日收在布林下軌之下）
nyse_price = "US$374.67（7/29收）" # NYSE TSM 7/29 Wed收$374.67(-$17.64、-4.50% vs 7/28 $392.31、Yahoo API確認)——美股由分化轉全面補跌；SOX同日-5.33%收10,447.49自史高-28.6%、VIX +13.5%升破20至20.66、NVDA -3.55%翻黑；僅AAPL -0.56%抗跌
volume = "68,140（7/29收）"    # 台股7/29官方成交量68,140張（TWSE口徑、量增50%、達5日均量1.75倍、為7/17恐慌量97,362的70%——恐慌拋售升級）、成交額1,515.52億
news_summary = "報告日2026-07-30(Thu)。⚠️股災第2日：台股7/29大盤-1,564.18點(-3.76%)收40,039.18、兩日累計約-3,595點貼近4萬關口；2330官方收NT$2,200(-80、-3.51%、開2,260/高2,280/低2,180、量68,140張、額1,515.52億、675,528筆、TWSE確認)——13:00後恐慌殺盤觸2,180尾盤拉回、跌幅與大盤相當、抗跌角色弱化；供應鏈續殺：日月光-9.93%、華邦電-9.72%、聯電-9.69%。✅籌碼賣壓明顯收斂（最重要邊際變化）：外資賣超-11,211張(連6賣、自-14,659收斂24%、T86官方)、投信+1,652連4買放大6倍、自營+4,893連5買；全市場BFI82U賣超-351.5億（外資-222.5億）自-1,176億收斂70%——恐慌高峰或已過。⚠️美股7/29由分化轉全面補跌：TSM -4.50%收$374.67、SOX -5.33%收10,447.49自史高-28.6%——KLAC -10.80%、MU -9.94%、AMAT -8.40%、AMD -5.51%、NVDA -3.55%翻黑；僅AAPL -0.56%抗跌；VIX +13.5%升破20至20.66——前日分化格局終結。📅FOMC 7/29鷹派按兵不動：利率3.50-3.75%連5次、投票9-3分歧（3票主張升息）、主席Warsh未鬆口；WTI反彈~$84.2。⭐盤後財報第一根修復錨：MSFT營收$90.01B大超預期、Azure +43%、FY26 Azure首破$1,000億——盤後大漲+8.83%收$425.01（正規盤收$390.54、-0.71%；財報剛公布時僅約+3%、電話會後持續擴大）；META營收+28%超預期惟EPS miss ~14%、2026 CapEx $125-145B——盤後-7.45%收$542.00（正規盤收$585.61；一度殺至-9.64%）——AI CapEx未收縮、對TSMC訂單能見度正面。⚠️韓股7/29續挫：SK海力士財報創史高仍-9.61%、三星-5.23%；KOSPI 7/29官方收5,663.24(-5.98%、前收6,023.66、連2度熔斷——已補查證)。📈今晨(7/30)盤中09:26亞股同步回穩：台股大盤40,235.08(+0.49%、日內39,404.65-40,452.01)、2330開2,205/高2,240/低2,190於2,200上下震盤、KOSPI 5,665.17(+0.03%、區間5,547-5,781)、日經約+1.59%——屬極度超賣後技術性回穩、非趨勢反轉（盤中數據）。📊ADR溢價+10.2%（匯率32.36）。技術面(官方收盤計算)：RSI 35.6、MACD綠柱-47.8急擴且DIF -22.2深陷零軸下、KDJ死叉第4日且J -9.3轉負（極度超賣、歷史罕見）；支撐2,180/2,134(120MA)/2,100、壓力2,248(布林下軌)/2,260/2,280/2,317。⭐觀察：守2,180／收復2,248、今晚AAPL/AMZN財報、外資賣超收斂延續性、韓股止跌、232條款今明兩日出爐。"
change_color = "FF0000"        # 紅色（7/29官方-3.51%下跌）

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

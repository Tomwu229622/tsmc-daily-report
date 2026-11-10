"""
rotate_archive.py - 把產生器原始碼內超過工作集上限的舊條目搬進 archive JSON。
  IR_CALENDAR（升冪）保留最後 IR_KEEP 條，較舊者併入 history/ir_calendar_archive.json（升冪）
  EPS_REVISION["triggers"]（降冪，新的在前）保留前 TR_KEEP 條，其餘接在 history/eps_triggers_archive.json 後（降冪）
用法：python -X utf8 rotate_archive.py [--dir 產生器所在資料夾] [--backup-dir 備份資料夾] [--dry-run]
預設 dir = 本腳本所在資料夾；備份預設到 dir/backup/（請勿把 .bak 放進 git 工作樹，部署時改用 --backup-dir 指到暫存區）。
流程：ast 定位 → literal_eval 取值 → repr 重寫 → 寫回前 literal_eval 讀回比對原值完全相等 → 才落檔。
"""
import ast, io, json, os, sys, shutil, argparse, unicodedata
from datetime import date

IR_KEEP = 30
TR_KEEP = 15
GEN = "generate_tsmc_html_report.py"


def load_json(fp):
    if not os.path.exists(fp):
        return []
    with open(fp, "r", encoding="utf-8") as f:
        return [tuple(x) for x in json.load(f)]


def save_json(fp, rows):
    os.makedirs(os.path.dirname(fp), exist_ok=True)
    with open(fp, "w", encoding="utf-8", newline="\n") as f:
        json.dump([list(r) for r in rows], f, ensure_ascii=False, indent=1)
        f.write("\n")


def locate(tree):
    ir = eps = None
    for n in tree.body:
        if isinstance(n, ast.Assign) and len(n.targets) == 1 and isinstance(n.targets[0], ast.Name):
            if n.targets[0].id == "IR_CALENDAR":
                ir = n
            elif n.targets[0].id == "EPS_REVISION":
                eps = n
    assert ir is not None and eps is not None, "找不到 IR_CALENDAR / EPS_REVISION"
    trig = None
    for k, v in zip(eps.value.keys, eps.value.values):
        if isinstance(k, ast.Constant) and k.value == "triggers":
            trig = v
    assert trig is not None, "找不到 triggers"
    return ir, trig


def lit(x):
    """字串字面值：用 repr，但把 repr 對不可列印字元（如 U+FE0F）產生的 uXXXX 跳脫還原成原字元。"""
    if not isinstance(x, str):
        return repr(x)
    r = repr(x)
    out, i = [], 0
    while i < len(r):
        c = r[i]
        if c == "\\" and i + 1 < len(r):
            n = r[i + 1]
            width = {"u": 4, "x": 2, "U": 8}.get(n)
            if width:
                ch = chr(int(r[i + 2:i + 2 + width], 16))
                if unicodedata.category(ch) not in ("Cc", "Cs", "Co", "Cn", "Zl", "Zp"):
                    out.append(ch)
                    i += 2 + width
                    continue
            out.append(c + n)
            i += 2
            continue
        out.append(c)
        i += 1
    return "".join(out)


def fmt_tuple(r):
    return "(" + ", ".join(lit(x) for x in r) + ")"


def fmt_list(rows, indent):
    pad = " " * indent
    return "[\n" + "".join(pad + fmt_tuple(r) + ",\n" for r in rows)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dir", default=os.path.dirname(os.path.abspath(__file__)))
    ap.add_argument("--backup-dir", default=None)
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    gp = os.path.join(a.dir, GEN)
    src = io.open(gp, encoding="utf-8", newline="").read()
    assert "\r" not in src
    lines = src.split("\n")
    tree = ast.parse(src)
    ir_node, trig_node = locate(tree)
    ir_all = [tuple(x) for x in ast.literal_eval(ir_node.value)]
    tr_all = [tuple(x) for x in ast.literal_eval(trig_node)]
    print("before: IR=%d triggers=%d" % (len(ir_all), len(tr_all)))
    ir_move, ir_keep = ir_all[:-IR_KEEP], ir_all[-IR_KEEP:]
    tr_keep, tr_move = tr_all[:TR_KEEP], tr_all[TR_KEEP:]
    if not ir_move and not tr_move:
        print("nothing to rotate")
        return
    hist = os.path.join(a.dir, "history")
    ir_arc = load_json(os.path.join(hist, "ir_calendar_archive.json"))
    tr_arc = load_json(os.path.join(hist, "eps_triggers_archive.json"))
    ir_keys = {(r[0], r[1]) for r in ir_arc}
    new_ir_arc = ir_arc + [r for r in ir_move if (r[0], r[1]) not in ir_keys]
    new_tr_arc = list(tr_move) + [t for t in tr_arc if t not in set(tr_move)]
    # 重寫原始碼：先改後面的區段（triggers），再改 IR，行號才不位移
    # 區段範圍 = 該 list 節點的 lineno..end_lineno，這些行必須恰好是 "[" 開頭到 "]" 結尾的獨立行
    def span(node):
        return node.lineno - 1, node.end_lineno  # slice 用

    edits = []
    # IR_CALENDAR：整個 list 在 "IR_CALENDAR = [" 起、"]" 止
    i0, i1 = span(ir_node)
    assert lines[i0].startswith("IR_CALENDAR = [") and lines[i1 - 1] == "]", "IR_CALENDAR 行距格式不符"
    new_ir_lines = ("IR_CALENDAR = " + fmt_list(ir_keep, 4) + "]").split("\n")
    edits.append((i0, i1, new_ir_lines, "IR"))
    # triggers：list 節點起訖行
    t0, t1 = trig_node.lineno - 1, trig_node.end_lineno
    assert lines[t0].rstrip().endswith("[") and lines[t1 - 1].strip() == "],", "triggers 行距格式不符"
    new_tr_lines = (fmt_list(tr_keep, 8)).split("\n")
    new_tr_lines[0] = lines[t0]  # 保留 '    "triggers": [' 原樣
    new_tr_lines = new_tr_lines[:-1] + [lines[t1 - 1]]  # 去掉 fmt 尾端空字串，補原結尾行
    edits.append((t0, t1, new_tr_lines, "TR"))
    for st, en, nl, _ in sorted(edits, key=lambda e: -e[0]):
        lines[st:en] = nl
    new_src = "\n".join(lines)
    # 讀回驗證
    tree2 = ast.parse(new_src)
    ir2, trig2 = locate(tree2)
    ir2v = [tuple(x) for x in ast.literal_eval(ir2.value)]
    tr2v = [tuple(x) for x in ast.literal_eval(trig2)]
    assert ir2v == ir_keep, "IR 讀回不相等"
    assert tr2v == tr_keep, "triggers 讀回不相等"
    # 合併後與原完整清單相等
    merged_ir = [r for r in new_ir_arc if (r[0], r[1]) not in {(x[0], x[1]) for x in ir2v}] + ir2v
    merged_tr = tr2v + [t for t in new_tr_arc if t not in set(tr2v)]
    assert merged_ir == [r for r in ir_arc if (r[0], r[1]) not in {(x[0], x[1]) for x in ir_all}] + ir_all, "IR 合併後不等於原完整清單"
    assert merged_tr == tr_all + [t for t in tr_arc if t not in set(tr_all)], "triggers 合併後不等於原完整清單"
    # JSON 讀回驗證
    assert "\\u" not in json.dumps([list(r) for r in new_ir_arc + new_tr_arc], ensure_ascii=False)
    if a.dry_run:
        print("dry-run OK: IR keep %d archive %d; triggers keep %d archive %d" % (len(ir_keep), len(new_ir_arc), len(tr_keep), len(new_tr_arc)))
        return
    bdir = a.backup_dir or os.path.join(a.dir, "backup")
    os.makedirs(bdir, exist_ok=True)
    shutil.copy(gp, os.path.join(bdir, GEN + ".bak-" + date.today().strftime("%Y%m%d")))
    save_json(os.path.join(hist, "ir_calendar_archive.json"), new_ir_arc)
    save_json(os.path.join(hist, "eps_triggers_archive.json"), new_tr_arc)
    with open(gp, "w", encoding="utf-8", newline="") as f:
        f.write(new_src)
    # 落檔後再讀一次
    again = ast.literal_eval(locate(ast.parse(io.open(gp, encoding="utf-8").read()))[0].value)
    assert [tuple(x) for x in again] == ir_keep
    print("rotated: IR keep %d / archive %d; triggers keep %d / archive %d" % (len(ir_keep), len(new_ir_arc), len(tr_keep), len(new_tr_arc)))


if __name__ == "__main__":
    main()

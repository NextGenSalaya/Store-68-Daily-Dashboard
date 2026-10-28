"""Topic 2 of the Store 68 Daily Board: KVI / Own Brand / Makro Mail stock monitoring, pack & component,
and O2O item shortage. Reads the five source files from anywhere under this folder and writes JSON datasets.

Names of customers, salespeople (SE), pickers and member numbers are never written out: those sheets are
only summed up.
Usage: python build_topic2.py [out_folder]
"""
import glob, json, os, re, sys, warnings
from collections import defaultdict, Counter
from datetime import datetime, date
import openpyxl

warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "data_t2")
os.makedirs(OUT, exist_ok=True)


def find(pattern):
    hits = sorted(glob.glob(os.path.join(HERE, "**", pattern), recursive=True), key=os.path.getmtime)
    if not hits:
        sys.exit(f"{pattern} not found under {HERE}")
    print("reading", hits[-1])
    return hits[-1]


def num(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(",", "").replace("%", "").replace("~", "").strip()
    try:
        return float(s)
    except ValueError:
        return None


def rnd(v, d=2):
    v = num(v)
    return None if v is None else round(v, d)


def txt(v):
    return "" if v is None else str(v).strip()


def code(v):
    n = num(v)
    return str(int(n)) if n is not None and n == int(n) else txt(v)


def day(v):
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    m = re.match(r"(\d{2})[-_/](\w{2,3})[-_/](\d{4})", txt(v))
    if m and not m.group(2).isdigit():
        try:
            return datetime.strptime(f"{m.group(1)}-{m.group(2)}-{m.group(3)}", "%d-%b-%Y").strftime("%Y-%m-%d")
        except ValueError:
            return None
    m = re.match(r"(\d{2})-(\d{2})-(\d{4})", txt(v))
    return f"{m.group(3)}-{m.group(2)}-{m.group(1)}" if m else None


def save(name, rows):
    with open(os.path.join(OUT, name + ".json"), "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, separators=(",", ":"))
    print(f"{name}: {len(rows)} rows")


DIV = {"FF": "FRESH FOOD", "DF": "DRY FOOD", "NF": "NON FOOD"}

# ------------------------------------------------------------------ stock monitoring item lists
items, meta = [], []
for lst, pattern in (("KVI", "1.KVI Stock Monitoring*.xlsx"), ("Own Brand", "2.OB Stock Monitoring*.xlsx"),
                     ("Makro Mail", "3.Mmail Stock Monitoring*.xlsx")):
    wb = openpyxl.load_workbook(find(pattern), read_only=True, data_only=True)
    for sh in ("FF", "DF", "NF"):
        ws = wb[sh]
        head, started, blanks = {}, False, 0
        for r in ws.iter_rows(max_col=13, values_only=True):
            if not started:
                if txt(r[8]):
                    head[txt(r[8])] = r[10]
                if txt(r[0]).startswith("Mmail #") and lst == "Makro Mail" and not any(m.get("period") for m in meta):
                    meta.append({"key": "Makro Mail period", "period": txt(r[0])})
                started = txt(r[1]) == "Art No"
                continue
            if r[1] is None:
                blanks += 1
                if blanks > 20:
                    break
                continue
            blanks = 0
            stock, oo = num(r[4]), num(r[6]) or 0
            status = "มีของ" if (stock or 0) > 0 else ("ขาด มี order" if oo > 0 else "ขาด ไม่มี order")
            items.append({"key": f"{lst}|{code(r[1])}", "list": lst, "division": DIV[sh], "dept": code(r[0]), "item": code(r[1]),
                          "desc": txt(r[2]), "dms": rnd(r[3], 2), "stock": rnd(stock, 1), "stock_day": rnd(r[5], 0), "on_order": rnd(oo, 0),
                          "reorder": txt(r[7]) if txt(r[7]) not in ("0", "") else "", "last_sale": day(r[8]), "last_del": day(r[9]),
                          "item_status": txt(r[10]), "pack": rnd(r[11], 0), "status": status})
# keys must be unique: an item listed twice in one sheet keeps its first row
seen, uniq = set(), []
for it in items:
    if it["key"] not in seen:
        seen.add(it["key"])
        uniq.append(it)
save("t2_items", uniq)
save("t2_meta", meta)

# KVI price-survey list (KVI file)
wb = openpyxl.load_workbook(find("1.KVI Stock Monitoring*.xlsx"), read_only=True, data_only=True)
ws = wb["New Price Survey (KVI) Lists"]
kvi, upd = [], ""
for i, r in enumerate(ws.iter_rows(max_col=17, values_only=True)):
    if i == 1:
        upd = txt(r[0])
    if i < 3 or r[4] is None or num(r[4]) is None:
        continue
    kvi.append({"key": code(r[4]), "div": txt(r[0]), "buyer": txt(r[1]), "class": f"{code(r[2])} {txt(r[3])}", "item": code(r[4]),
                "name_th": txt(r[5]), "name_en": txt(r[6]), "location": txt(r[7]), "scope": txt(r[8]), "kvi_type": txt(r[9]), "remark": txt(r[10])})
seen, kv2 = set(), []
for k in kvi:
    if k["key"] not in seen:
        seen.add(k["key"])
        kv2.append(k)
save("t2_kvi_list", kv2)
meta.append({"key": "KVI list updated", "period": upd})
save("t2_meta", meta)

# ------------------------------------------------------------------ summary sheet (all items vs store-order items)
wb = openpyxl.load_workbook(find("4.ใบสรุป KVI-OB-MMail*.xlsx"), data_only=True)
ws = wb["Summary"]
summ = []
LISTS = {"MMail": "Makro Mail", "KVI": "KVI", "Own Brand": "Own Brand"}
for r in range(1, ws.max_row + 1):
    b = txt(ws.cell(r, 2).value)
    lst = next((v for k, v in LISTS.items() if b.startswith(k)), None)
    if not lst or "Monitoring" not in b and not b.startswith("Own Brand"):
        continue
    first = next(rr for rr in range(r, r + 12) if txt(ws.cell(rr, 2).value) == "Item in Stock")
    for view, cols in (("ทุกรายการ", (4, 5, 6, 7)), ("Store Order", (9, 10, 11, 12))):
        for c, div in zip(cols, ("FRESH FOOD", "DRY FOOD", "NON FOOD", "Total")):
            ins, nno, no, tot = (num(ws.cell(first + k, c).value) for k in range(4))
            summ.append({"key": f"{lst}|{view}|{div}", "list": lst, "view": view, "division": div, "in_stock": ins,
                         "oos_not_on_order": nno, "oos_on_order": no, "total": tot,
                         "service_level": round(ins / tot * 100, 2) if tot else None})
save("t2_summary", summ)

# pack & component
pk = []
for r in list(wb["Pack"].iter_rows(min_row=2, values_only=True)):
    if r[3] is None:
        continue
    ps, cs = num(r[6]) or 0, num(r[10]) or 0
    # packs are mostly kept at zero stock; what matters is whether the component behind the pack has stock
    flag = "Pack มีของ" if ps > 0 else ("Component มีของ" if cs > 0 else "Component หมด")
    pk.append({"key": f"{code(r[3])}|{code(r[8])}", "group": txt(r[0]), "dept": code(r[1]), "class": code(r[2]), "pack_item": code(r[3]),
               "pack_desc": txt(r[5]), "pack_stock": rnd(ps, 1), "pack_price": rnd(r[7]), "comp_item": code(r[8]), "comp_desc": txt(r[9]),
               "comp_stock": rnd(cs, 1), "comp_price": rnd(r[11]), "dms": rnd(r[12], 2), "flag": flag})
seen, pk2 = set(), []
for p in pk:
    if p["key"] not in seen:
        seen.add(p["key"])
        pk2.append(p)
save("t2_pack", pk2)

# ------------------------------------------------------------------ item shortage (O2O)
wb = openpyxl.load_workbook(find("3.%Item Shortage MTD*.xlsx"), read_only=True, data_only=True)
data_date = None

# order lines delivered short (no customer, SE or remark text)
lines = []
for i, r in enumerate(wb["Shortage _Reason"].iter_rows(max_col=14, values_only=True)):
    if i == 1 and isinstance(r[3], (datetime, date)):
        data_date = day(r[3])
    if i < 4 or r[0] is None or num(r[6]) is None:
        continue
    lines.append({"item": code(r[0]), "desc": txt(r[1]), "division": txt(r[3]), "ordered": rnd(r[4], 2), "shipped": rnd(r[5], 2),
                  "short": rnd(r[6], 2), "pct": rnd((num(r[7]) or 0) * 100, 1), "reason": txt(r[8]) or "ไม่ระบุ", "day": day(r[9])})
save("t2_short_lines", [dict(l, key=str(i)) for i, l in enumerate(lines)])

# daily shortage counts per item (Shortage By item): total per day, and per item with its daily counts
ws = wb["Shortage By item"]
rows = list(ws.iter_rows(values_only=True))
hdr_i = next(i for i, r in enumerate(rows) if txt(r[1]) == "Item No" and txt(r[2]) == "Item Description")
dates = [(j, day(v)) for j, v in enumerate(rows[hdr_i]) if isinstance(v, (datetime, date))]
per_day, per_item = Counter(), []
for r in rows[hdr_i + 1:]:
    if r[1] is None or not code(r[1]).isdigit():
        continue
    counts = {d: int(num(r[j]) or 0) for j, d in dates}
    for d, n in counts.items():
        per_day[d] += n
    per_item.append({"item": code(r[1]), "desc": txt(r[2]), "total": sum(counts.values()),
                     "last7": sum(n for d, n in counts.items() if d >= sorted(counts)[-7])})
save("t2_short_by_day", [{"day": d, "lines": per_day[d]} for _, d in dates])
seen, pi = set(), []
for p in sorted(per_item, key=lambda x: -x["total"]):
    if p["item"] not in seen:
        seen.add(p["item"])
        pi.append(p)
save("t2_short_items", pi)

# orders short by division and department (Shortage By Dept)
ws = wb["Shortage By Dept."]
by_div, by_dept = [], []
for r in list(ws.iter_rows(max_col=14, values_only=True))[3:]:
    if txt(r[4]) and num(r[5]) is not None and txt(r[4]) not in ("(blank)",):
        by_div.append({"division": txt(r[4]), "orders": rnd(r[5], 0)})
    if txt(r[7]) and num(r[8]) is not None:
        by_dept.append({"dept": txt(r[7]), "orders": rnd(r[8], 0)})
save("t2_short_div", by_div)
save("t2_short_dept", by_dept)

# extra detail from DataExport: by buyer group, order way and item status (no customer or SE names)
agg = defaultdict(lambda: {"lines": 0, "short": 0.0})
for i, r in enumerate(wb["DataExport"].iter_rows(max_col=40, values_only=True)):
    if i == 0 or r[15] is None:
        continue
    for dim, val in (("Buyer", txt(r[34]) or "ไม่ระบุ"), ("Order Way", txt(r[32]) or "ไม่ระบุ"), ("Item Status", txt(r[35]) or "ไม่ระบุ"),
                     ("Channel", txt(r[8]) or "ไม่ระบุ")):
        a = agg[(dim, val)]
        a["lines"] += 1
        a["short"] += num(r[22]) or 0
save("t2_short_breakdown", [{"key": f"{d}|{v}", "dim": d, "value": v, "lines": a["lines"], "short": round(a["short"], 1)}
                            for (d, v), a in sorted(agg.items(), key=lambda kv: (kv[0][0], -kv[1]["lines"]))])

# picking OOS history (Data OOS overview1): by month and reason, by department, top items (no worker names)
by_month, by_dept_r, by_item = Counter(), Counter(), defaultdict(lambda: {"count": 0, "short": 0.0})
for i, r in enumerate(wb["Data OOS overview1"].iter_rows(max_col=36, values_only=True)):
    if i == 0 or r[12] is None:
        continue
    d = day(r[9]) or ""
    reason = txt(r[22]) or "ไม่ระบุ"
    by_month[(d[:7], reason)] += 1
    by_dept_r[(txt(r[17]) or "ไม่ระบุ", txt(r[15]) or "", reason)] += 1
    it = by_item[(code(r[12]), txt(r[13]), txt(r[15]))]
    it["count"] += 1
    it["short"] += num(r[26]) or 0
save("t2_pick_month", [{"key": f"{m}|{rs}", "month": m, "reason": rs, "count": n} for (m, rs), n in sorted(by_month.items()) if m])
save("t2_pick_dept", [{"key": f"{d}|{rs}", "dept": d, "division": dv, "reason": rs, "count": n}
                      for (d, dv, rs), n in sorted(by_dept_r.items(), key=lambda kv: -kv[1])])
save("t2_pick_items", [{"key": it, "item": it, "desc": ds, "division": dv, "count": a["count"], "short": round(a["short"], 1)}
                       for (it, ds, dv), a in sorted(by_item.items(), key=lambda kv: -kv[1]["count"])[:200]])

# today's out-of-stock log at picking (Data oos): by division, reason, hour, status; and the item list (no names)
today, by_hour = [], Counter()
for i, r in enumerate(wb["Data oos"].iter_rows(max_col=40, values_only=True)):
    if i == 0 or r[15] is None:
        continue
    today.append({"key": txt(r[6]) or str(i), "hour": int(num(r[0]) or 0), "section": txt(r[2]), "division": txt(r[22]),
                  "reason": txt(r[13]) or "ไม่ระบุ", "status": txt(r[7]), "item": code(r[15]), "desc": txt(r[16]), "short": rnd(r[17], 2),
                  "plan": rnd(r[19], 2), "dept": txt(r[30]), "shelf": txt(r[26]), "channel": txt(r[24]), "slot": txt(r[23]),
                  "wait": txt(r[14]), "floor_reason": txt(r[39]) or txt(r[37])})
seen, td2 = set(), []
for t in today:
    if t["key"] not in seen:
        seen.add(t["key"])
        td2.append(t)
save("t2_oos_today", td2)

# O2O orders shipped (DataAO): per day and per payment method / delivery type
ao_day, ao_dim = defaultdict(lambda: {"orders": 0, "amount": 0.0}), defaultdict(lambda: {"orders": 0, "amount": 0.0})
for i, r in enumerate(wb["DataAO"].iter_rows(max_col=40, values_only=True)):
    if i == 0 or r[0] is None:
        continue
    d = day(r[11])
    amt = num(r[37]) or 0
    if d:
        ao_day[d]["orders"] += 1
        ao_day[d]["amount"] += amt
    for dim, val in (("Payment", txt(r[7]) or "ไม่ระบุ"), ("Delivery", txt(r[15]) or "ไม่ระบุ"), ("Channel", txt(r[17]) or "ไม่ระบุ"),
                     ("Order Type", txt(r[13]) or "ไม่ระบุ")):
        ao_dim[(dim, val)]["orders"] += 1
        ao_dim[(dim, val)]["amount"] += amt
save("t2_ao_day", [{"day": d, "orders": v["orders"], "amount": round(v["amount"])} for d, v in sorted(ao_day.items())])
save("t2_ao_dim", [{"key": f"{d}|{v}", "dim": d, "value": v, "orders": a["orders"], "amount": round(a["amount"])}
                   for (d, v), a in sorted(ao_dim.items(), key=lambda kv: (kv[0][0], -kv[1]["orders"]))])

meta.append({"key": "Shortage data date", "period": data_date or ""})
save("t2_meta", meta)

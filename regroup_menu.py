"""Rewrite the page menu: pages grouped by subject, and pages on the same subject merged into one menu item.
A menu item lists the sections it shows in data-page (space separated); its id is tab-<first section>."""
import html, re
P = "board/index.html"
t = open(P, encoding="utf-8").read()
start = t.index('<nav class="b68-nav" aria-label="Menu">')
end = t.index("</nav>", start)
# (menu id, sections shown, label)
GROUPS = [
    ("grp-sales", "Overall Trend", [
        ("sales", "sales", "ยอดขาย & ลูกค้า"),
        ("category", "category t5cat items", "Category, Class & สินค้า"),
        ("t5item", "t5item", "Sale by Item & Top Drop Customer"),
        ("t3ff", "t3ff t5top50", "Fresh Food Sale & Top 50"),
        ("t5hamper", "t5hamper t5soft", "Hamper & Soft-serve")]),
    ("grp-stock", "Stock & Order", [
        ("stock", "stock t5stock", "Stock & Stock Days"),
        ("t5mu", "t5mu t5watch", "Stock ใกล้หมด & ติดลบ"),
        ("t4over", "t4over t4floor", "Over Stock & On Order"),
        ("t5aging", "t5aging t5gone", "Aging, NBS & Clearance"),
        ("t5repl", "t5repl", "Central Replenishment"),
        ("t5top", "t5top t3il", "Top Items & Item List")]),
    ("grp-svc", "Service Level · KVI", [
        ("service", "service t2sum", "Service level"),
        ("t2kvi", "t2items", "KVI · OB · Makro Mail Monitoring"),
        ("t2survey", "t2survey", "KVI Price Survey List"),
        ("t2pack", "t2pack", "Pack & Component")]),
    ("grp-shk", "Shrinkage", [
        ("shrink", "shrink t3shk", "Shrinkage ภาพรวม & แผนก"),
        ("t3c20", "t3c20 t5shk", "Shrinkage by Code")]),
    ("grp-o2o", "O2O", [
        ("o2o", "o2o t2ao", "O2O ภาพรวม & Orders"),
        ("t3top", "t3top", "Focus Top 1000 O2O"),
        ("t2short", "t2short t2pick", "Item Shortage & OOS")]),
]
sections = set(re.findall(r'<section class="b68-page" id="(\w+)"', t)) | set(re.findall(r"\n    (t5\w+): \{", t))
used = [p for _, _, items in GROUPS for _, pages, _ in items for p in pages.split()]
assert sorted(used) == sorted(sections), (sections - set(used), set(used) - sections)
out = ['<nav class="b68-nav" aria-label="Menu">']
for gid, name, items in GROUPS:
    out.append(f'''      <div class="b68-group">
        <button type="button" class="b68-grp" aria-expanded="true" aria-controls="{gid}">
          <span>{html.escape(name)}</span>
          <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 15l6-6 6 6"/></svg>
        </button>
        <div class="b68-menu" id="{gid}">''')
    out += [f'          <button type="button" id="tab-{i}" data-page="{p}">{html.escape(label)}</button>' for i, p, label in items]
    out.append("        </div>\n      </div>")
t = t[:start] + "\n".join(out) + "\n    " + t[end:]
open(P, "w", encoding="utf-8").write(t)
print("ok", len(used), "sections in", sum(len(i) for _, _, i in GROUPS), "menu items")

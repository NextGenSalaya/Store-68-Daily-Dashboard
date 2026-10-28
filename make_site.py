"""Build the public website version of the Store 68 Daily Board into docs/ (GitHub Pages serves this folder).

docs/index.html = board/index.html plus a small runtime that does what the claude.ai dashboard does for the page:
loads each dataset from docs/data/<id>.json the first time the page asks for it, runs the page's summary calcs,
keeps the Date Period / Division filters, and puts the open menu item in the address (#link).
docs/data/ = every JSON file from data_ad, data_t2 .. data_t5 (run the build_*.py extractors first).
Usage: python make_site.py
"""
import glob, os, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
SITE = os.path.join(HERE, "docs")
DATA = os.path.join(SITE, "data")
os.makedirs(DATA, exist_ok=True)
for f in glob.glob(os.path.join(DATA, "*.json")):
    os.remove(f)
n = 0
for d in ("data_ad", "data_t2", "data_t3", "data_t4", "data_t5"):
    for f in glob.glob(os.path.join(HERE, d, "*.json")):
        shutil.copy2(f, DATA)
        n += 1

RUNTIME = r"""<!doctype html>
<html lang="th">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Store 68 Daily Board</title>
<style>
:root { --color-fg: #1b1f24; --color-bg: #ffffff; --color-panel: #ffffff; --color-border-line: #dde2ea; --color-fg-muted: #5b6472;
  --color-bad: #d4343b; --cds-chart-grid: #e6e9ef; --cds-chart-axis: #9aa3b0; --cds-surface-popover: #ffffff; --cds-border: #dde2ea;
  --cds-radius: 6px; --cds-shadow-popover: 0 4px 16px rgb(15 40 90 / .15); --cds-pad-xs: 6px; --cds-pad-sm: 10px;
  --cds-font-size-caption: 13px; --cds-font-size-body: 15px; --cds-text-secondary: #5b6472; --cds-dur-slow: .3s;
  --font-anthropic-sans: system-ui, sans-serif; color-scheme: light; }
html, body { margin: 0; background: #eef2f8; }
body { padding: 0; }
.site-loading { position: fixed; right: 14px; bottom: 14px; background: #0b3a8a; color: #fff; font: 13px system-ui, sans-serif;
  padding: 6px 12px; border-radius: 999px; box-shadow: 0 2px 8px rgb(0 0 0 / .2); z-index: 50; }
.site-loading[hidden] { display: none; }
</style>
<script src="https://cdnjs.cloudflare.com/ajax/libs/d3/7.8.5/d3.min.js"></script>
<script>
// Stand-in for the claude.ai dashboard runtime: datasets load from data/<id>.json on first use.
(() => {
  const store = new Map();        // id -> {status, data}
  const calcs = new Map();        // id -> {inputs, fn}
  const cache = new Map();        // calc id -> {stamp, result}
  const subs = [];
  const DEF = { period: 'MTD', div: 'ALL' };
  let params = { ...DEF };
  let stamp = 0, pending = 0, queued = false;
  const badge = () => document.querySelector('.site-loading');
  const notify = () => {
    if (queued) return;
    queued = true;
    requestAnimationFrame(() => { queued = false; subs.forEach((f) => f()); });
  };
  const load = (id) => {
    if (store.has(id)) return store.get(id);
    const s = { status: 'loading', data: undefined };
    store.set(id, s);
    pending += 1; if (badge()) badge().hidden = false;
    fetch(`data/${encodeURIComponent(id)}.json`, { cache: 'no-cache' })
      .then((r) => { if (!r.ok) throw new Error(r.status); return r.json(); })
      .then((d) => { s.status = 'ok'; s.data = d; })
      .catch(() => { s.status = 'error'; })
      .finally(() => { pending -= 1; stamp += 1; if (!pending && badge()) badge().hidden = true; notify(); });
    return s;
  };
  const runCalc = (id) => {
    const c = calcs.get(id);
    const ins = c.inputs.map(load);
    if (ins.some((s) => s.status === 'error')) return { status: 'error' };
    if (ins.some((s) => s.status !== 'ok')) return { status: 'loading' };
    const hit = cache.get(id);
    if (hit && hit.stamp === stamp) return hit.result;
    let result;
    try { result = { status: 'ok', data: c.fn(...ins.map((s) => s.data)) }; } catch (e) { console.error(id, e); result = { status: 'error' }; }
    cache.set(id, { stamp, result });
    return result;
  };
  window.dash = {
    data: (id) => (calcs.has(id) ? runCalc(id) : load(id)),
    calc: (id, o) => { calcs.set(id, o); },
    params: () => params,
    setParams: (p) => { params = { ...params, ...p }; notify(); },
    resetParams: () => { params = { ...DEF }; notify(); },
    setLink: (l) => { history.replaceState(null, '', '#' + l); },
    onData: (f) => { subs.push(f); notify(); },
  };
  // the menu item in the address bar opens when the link changes
  window.addEventListener('hashchange', () => {
    const b = document.getElementById('tab-' + location.hash.slice(1));
    if (b) b.click();
  });
})();
</script>
</head>
<body>
<div class="site-loading" hidden>กำลังโหลดข้อมูล…</div>
"""
page = open(os.path.join(HERE, "board", "index.html"), encoding="utf-8").read()
open(os.path.join(SITE, "index.html"), "w", encoding="utf-8").write(RUNTIME + page + "\n</body>\n</html>\n")
open(os.path.join(SITE, ".nojekyll"), "w").close()
print(f"docs/index.html written; {n} data files in docs/data")

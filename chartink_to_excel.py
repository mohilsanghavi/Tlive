import json, re, sys, time, pathlib, datetime
from zoneinfo import ZoneInfo
from playwright.sync_api import sync_playwright
from openpyxl import Workbook

ROOT = pathlib.Path(__file__).resolve().parent
cfg = json.loads((ROOT / "groups.json").read_text())
URL = cfg["scan_url"]
groups = [g for g in cfg["groups"] if g.get("on")]
if not groups:
    sys.exit("No groups are switched on in groups.json")


def cond_of(sc):
    m = re.match(r"^\s*\(\s*(?:\{[^}]*\}\s*)?\(([\s\S]*)\)\s*\)\s*$", sc or "")
    return re.sub(r"\s+", " ", m.group(1)).strip() if m else None


with sync_playwright() as p:
    browser = p.chromium.launch()
    ctx = browser.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36")
    page = ctx.new_page()
    got = {}

    def on_req(req):
        if "/screener/process" in req.url and req.method == "POST" and req.post_data:
            try:
                d = json.loads(req.post_data)
                if "column_clause" in d:
                    got.update(d)
            except Exception:
                pass

    page.on("request", on_req)
    page.goto(URL, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_selector("text=Run Scan", timeout=60000)
    page.get_by_role("button", name=re.compile("Run Scan")).first.click()
    for _ in range(60):
        if got:
            break
        time.sleep(0.5)
    if not got:
        sys.exit("Could not read the scan's column settings (Chartink may have blocked this server).")
    page.wait_for_timeout(3000)
    labels = page.evaluate("""() => {
      const t=[...document.querySelectorAll('table')].find(x=>/Stock Name/.test(x.innerText));
      return t?[...t.querySelectorAll('th')].map(h=>h.innerText.split('\\n')[0].trim()).slice(3):[];
    }""")
    token = page.evaluate("() => (document.querySelector('meta[name=\"csrf-token\"]')||{}).content")
    cond = cfg.get("filter_override") or cond_of(got["scan_clause"]) or "latest close > 0.01"
    cc = got["column_clause"]

    now = datetime.datetime.now(ZoneInfo("Asia/Kolkata"))
    summary = [["Scan link", URL, ""], ["Condition used", cond, ""],
               ["Downloaded (IST)", now.strftime("%Y-%m-%d %H:%M"), ""], [], ["Group", "Group ID", "Stocks"]]
    sheets = []
    for g in groups:
        clause = f"( {{{g['id']}}} ( {cond} ) )" if g["id"] else f"( {cond} )"
        res = page.evaluate("""async ([tok, body]) => {
          const r = await fetch('/screener/process',{method:'POST',headers:{'X-CSRF-TOKEN':tok,'Content-Type':'application/json','X-Requested-With':'XMLHttpRequest','Accept':'application/json'},body:JSON.stringify(body)});
          let j=null; try{ j=await r.json(); }catch(e){}
          return {status:r.status, data: j&&j.data};
        }""", [token, {"scan_clause": clause, "debug_clause": "", "column_clause": cc}])
        if res["data"] is None:
            sys.exit(f"Chartink refused group '{g['name']}' (HTTP {res['status']}).")
        data = res["data"]
        ks = [k for k in (data[0].keys() if data else []) if k.startswith("scan-column-")]
        hdr = ["Sr.", "Stock Name", "Symbol"] + ([(labels[i] if i < len(labels) and labels[i] else k) for i, k in enumerate(ks)] if ks else labels)
        rows = [[d.get("sr"), d.get("name"), d.get("nsecode")] + [d.get(k) for k in ks] for d in data]
        sheets.append((re.sub(r"[\\/?*\[\]:]", "-", g["name"])[:31], [hdr] + rows))
        summary.append([g["name"], g["id"] or "(default)", len(rows)])
        print(g["name"], len(rows))
        time.sleep(0.5)
    browser.close()

def sheet_json(rows):
    """rows -> list of {header: value} like the suite would read from the Excel sheet (blank rows skipped, blanks as "")."""
    if not rows:
        return []
    seen, hdr = {}, []
    for h in rows[0]:
        h = "" if h is None else str(h)
        if h in seen:
            seen[h] += 1
            h = f"{h}_{seen[h]}"
        else:
            seen[h] = 0
        hdr.append(h)
    out = []
    for r in rows[1:]:
        if all(v is None or v == "" for v in r):
            continue
        out.append({h: ("" if (i >= len(r) or r[i] is None) else r[i]) for i, h in enumerate(hdr)})
    return out


wb = Workbook()
ws = wb.active
ws.title = "Summary"
for r in summary:
    ws.append(r)
for name, rows in sheets:
    w = wb.create_sheet(name)
    for r in rows:
        w.append(r)
out = ROOT / "data"
out.mkdir(exist_ok=True)
KEEP_DAYS = 2   # change this number to keep more or fewer days
wb.save(out / f"Chartink_Groups_{now.strftime('%Y-%m-%d')}.xlsx")
# Fast copy for the suite: the same data as JSON (the suite reads this instead of opening the Excel)
fast = {"v": 1, "stamp": now.strftime("%Y-%m-%d %H:%M"),
        "sheets": {"Summary": {"aoa": summary}}}
for name, rows in sheets:
    fast["sheets"][name] = {"rows": sheet_json(rows)}
fast["names"] = ["Summary"] + [n for n, _ in sheets]
(out / f"Chartink_Groups_{now.strftime('%Y-%m-%d')}.json").write_text(json.dumps(fast, separators=(",", ":"), allow_nan=False), encoding="utf-8")
files = sorted(out.glob("Chartink_Groups_????-??-??.xlsx"))
for old_file in files[:-KEEP_DAYS]:
    old_file.unlink()
    old_file.with_suffix(".json").unlink(missing_ok=True)
print("saved", now.strftime("%Y-%m-%d %H:%M"), "| kept", [f.name for f in files[-KEEP_DAYS:]])

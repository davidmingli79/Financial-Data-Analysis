#!/usr/bin/env python3
import argparse, csv, datetime as dt, gzip, json, os, re, ssl, sys, time, urllib.parse, urllib.request
from typing import Dict, List, Optional, Tuple
try:
    from openpyxl import Workbook
except Exception:
    Workbook = None

UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36"

def iso(d: dt.date) -> str: return d.strftime("%Y-%m-%d")
def parse_iso(s: str) -> dt.date: return dt.datetime.strptime(s, "%Y-%m-%d").date()
def daterange(start: dt.date, end: dt.date):
    cur = start
    while cur <= end:
        yield cur
        cur += dt.timedelta(days=1)
def yyyymmdd(d: dt.date) -> str: return d.strftime("%Y%m%d")
def ensure_dir(p: str): os.makedirs(p, exist_ok=True)

def fetch_url(url: str, method: str="GET", data: Optional[bytes]=None, headers: Optional[Dict[str,str]]=None, timeout: int=30, retries: int=3) -> Tuple[int, Dict[str,str], bytes]:
    hdr = {"User-Agent": UA, "Accept": "*/*"}
    if headers: hdr.update(headers)
    last_err = None
    for i in range(retries):
        try:
            req = urllib.request.Request(url, data=data, headers=hdr, method=method)
            ctx = ssl.create_default_context()
            with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
                status = getattr(resp, "status", 200)
                resp_headers = {k: v for k, v in resp.headers.items()}
                body = resp.read()
            if resp_headers.get("Content-Encoding","").lower()=="gzip":
                try: body = gzip.decompress(body)
                except Exception: pass
            return status, resp_headers, body
        except Exception as e:
            last_err = e
            time.sleep(0.8*(2**i))
    raise RuntimeError(f"fetch failed after retries: {url} :: {last_err}")

def to_float(x) -> Optional[float]:
    if x is None: return None
    s = str(x).strip()
    if not s or s in (".","NA","N/A","-","null"): return None
    s = s.replace(",","")
    try: return float(s)
    except Exception: return None

def write_csv(path: str, rows: List[Tuple[str, Optional[float]]]):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f); w.writerow(["date","value"])
        for d,v in rows: w.writerow([d, "" if v is None else v])

def read_manual_csv(path: str) -> List[Tuple[str, Optional[float]]]:
    rows=[]
    with open(path,"r",encoding="utf-8") as f:
        r=csv.DictReader(f)
        for row in r:
            d=(row.get("date") or "").strip()
            if re.match(r"^\d{4}-\d{2}-\d{2}$", d):
                rows.append((d, to_float(row.get("value"))))
    rows.sort(key=lambda x:x[0])
    return rows

# ----- Providers -----
def fetch_stooq(symbol: str, start: dt.date, end: dt.date, interval: str="d") -> List[Tuple[str, Optional[float]]]:
    url = f"https://stooq.com/q/d/l/?s={urllib.parse.quote(symbol)}&d1={yyyymmdd(start)}&d2={yyyymmdd(end)}&i={interval}"
    status, headers, body = fetch_url(url)
    if status != 200: raise RuntimeError(f"Stooq HTTP {status}: {url}")
    text = body.decode("utf-8", errors="replace")
    lines = [ln for ln in text.splitlines() if ln.strip()]
    if len(lines)<=1: return []
    header = lines[0].split(",")
    i_date = next((i for i,h in enumerate(header) if h.lower()=="date"), 0)
    i_close = next((i for i,h in enumerate(header) if h.lower()=="close"), 4)
    out=[]
    for ln in lines[1:]:
        cols=ln.split(",")
        if len(cols)<=max(i_date,i_close): continue
        d=cols[i_date]
        if re.match(r"^\d{4}-\d{2}-\d{2}$", d):
            out.append((d, to_float(cols[i_close])))
    out.sort(key=lambda x:x[0])
    return out

def fetch_hkma(path: str, field: str, start: dt.date, end: dt.date, offset: int=0) -> List[Tuple[str, Optional[float]]]:
    url = f"https://api.hkma.gov.hk/public/{path}?offset={offset}"
    status, headers, body = fetch_url(url)
    if status != 200: raise RuntimeError(f"HKMA HTTP {status}: {url}")
    j = json.loads(body.decode("utf-8", errors="replace"))
    recs = (((j or {}).get("result") or {}).get("records")) or []
    out=[]
    s0,s1=iso(start), iso(end)
    for r in recs:
        d=r.get("end_of_day")
        if not d or d<s0 or d>s1: continue
        out.append((d, to_float(r.get(field))))
    out.sort(key=lambda x:x[0])
    return out

def csd_period_to_date(period: str) -> str:
    p=str(period).strip()
    if re.match(r"^\d{4}-\d{2}-\d{2}$", p): return p
    if re.match(r"^\d{6}$", p): return f"{p[:4]}-{p[4:6]}-01"
    if re.match(r"^\d{4}$", p): return f"{p}-01-01"
    return p

def fetch_csd(tableId: str, start: dt.date, end: dt.date, lang: str="en", full_series: int=1, param: str="") -> List[Tuple[str, Optional[float]]]:
    url = f"https://www.censtatd.gov.hk/api/get.php?id={tableId}&lang={lang}"
    if param: url += f"&param={urllib.parse.quote(param)}"
    else: url += f"&full_series={full_series}"
    status, headers, body = fetch_url(url)
    if status != 200: raise RuntimeError(f"C&SD HTTP {status}: {url}")
    j = json.loads(body.decode("utf-8", errors="replace"))
    dataset = j.get("dataSet") or j.get("dataset") or []
    if not isinstance(dataset, list): dataset=[]
    ignore={"period","figure","sd_value","count","started","finished","durationSeconds"}
    groups={}
    for rec in dataset:
        if not isinstance(rec, dict): continue
        dims=tuple(sorted([(k,str(v)) for k,v in rec.items() if k not in ignore]))
        groups.setdefault(dims, []).append(rec)
    chosen=max(groups.values(), key=lambda xs: len(xs), default=[])
    s0,s1=iso(start), iso(end)
    out=[]
    for rec in chosen:
        d=csd_period_to_date(rec.get("period"))
        if not d or len(d)!=10: continue
        if d<s0 or d>s1: continue
        out.append((d, to_float(rec.get("figure"))))
    out.sort(key=lambda x:x[0])
    return out

def bls_period_to_date(year: str, period: str) -> str:
    if period.startswith("M") and len(period)==3:
        m=int(period[1:])
        if 1<=m<=12: return f"{year}-{m:02d}-01"
    return f"{year}-01-01"

def fetch_bls_v1(series_id: str, start_year: int, end_year: int) -> List[Tuple[str, Optional[float]]]:
    url = "https://api.bls.gov/publicAPI/v1/timeseries/data/"
    payload = json.dumps({"seriesid":[series_id],"startyear":str(start_year),"endyear":str(end_year)}).encode("utf-8")
    status, headers, body = fetch_url(url, method="POST", data=payload, headers={"Content-Type":"application/json"})
    if status != 200: raise RuntimeError(f"BLS HTTP {status}: {series_id}")
    j = json.loads(body.decode("utf-8", errors="replace"))
    series_list = (((j.get("Results") or {}).get("series")) or [])
    if not series_list: return []
    data = series_list[0].get("data") or []
    out=[]
    for item in data:
        year=item.get("year"); period=item.get("period")
        if not year or not period or period=="M13": continue
        out.append((bls_period_to_date(year, period), to_float(item.get("value"))))
    out.sort(key=lambda x:x[0])
    return out

# NBS RSS
NBS_RSS = "https://www.stats.gov.cn/english/PressRelease/rss.xml"

def fetch_nbs_rss() -> List[Tuple[str,str,str]]:
    status, headers, body = fetch_url(NBS_RSS)
    if status != 200: raise RuntimeError(f"NBS RSS HTTP {status}")
    xml = body.decode("utf-8", errors="replace")
    items = re.findall(r"<item>(.*?)</item>", xml, flags=re.S|re.I)
    out=[]
    for it in items:
        title = re.search(r"<title><!\\[CDATA\\[(.*?)\\]\\]></title>", it, flags=re.S)
        if not title: title = re.search(r"<title>(.*?)</title>", it, flags=re.S)
        link = re.search(r"<link>(.*?)</link>", it, flags=re.S)
        pub = re.search(r"<pubDate>(.*?)</pubDate>", it, flags=re.S)
        if not title or not link or not pub: 
            continue
        t = re.sub(r"\\s+"," ", title.group(1)).strip()
        l = link.group(1).strip()
        p = pub.group(1).strip()
        # RFC822: "Tue, 12 May 2026 09:30:00 GMT"
        dtp=None
        for fmt in ("%a, %d %b %Y %H:%M:%S", "%a, %d %b %Y %H:%M:%S %Z"):
            try:
                dtp = dt.datetime.strptime(p[:25], "%a, %d %b %Y %H:%M:%S").date()
                break
            except Exception:
                pass
        if dtp is None:
            try: dtp = dt.datetime.strptime(p[:16], "%Y-%m-%d %H:%M").date()
            except Exception: continue
        out.append((iso(dtp), t, l))
    return out

def month_from_title(title: str) -> Optional[str]:
    m = re.search(r" in ([A-Za-z]+) (\\d{4})", title)
    if not m: return None
    month_name = m.group(1).lower()
    year = int(m.group(2))
    months = {"january":1,"february":2,"march":3,"april":4,"may":5,"june":6,"july":7,"august":8,"september":9,"october":10,"november":11,"december":12}
    mm = months.get(month_name)
    if not mm: return None
    return f"{year:04d}-{mm:02d}-01"

def quarter_from_title(title: str) -> Optional[str]:
    m = re.search(r"(First|Second|Third|Fourth) Quarter of (\\d{4})", title, flags=re.I)
    if not m: return None
    qmap = {"first":1,"second":2,"third":3,"fourth":4}
    q = qmap.get(m.group(1).lower()); y=int(m.group(2))
    if not q: return None
    month=(q-1)*3+1
    return f"{y:04d}-{month:02d}-01"

def strip_tags(html: str) -> str:
    html = re.sub(r"<script.*?</script>", " ", html, flags=re.S|re.I)
    html = re.sub(r"<style.*?</style>", " ", html, flags=re.S|re.I)
    txt = re.sub(r"<[^>]+>", " ", html)
    return re.sub(r"\\s+"," ", txt).strip()

def extract_cpi_yoy(txt: str) -> Optional[float]:
    m = re.search(r"CPI\\) (?:increased|decreased) by ([\\d\\.\\-]+)% year on year", txt, flags=re.I)
    return to_float(m.group(1)) if m else None

def extract_ppi_yoy(txt: str) -> Optional[float]:
    m = re.search(r"PPI\\) (?:increased|decreased) by ([\\d\\.\\-]+)% year on year", txt, flags=re.I)
    return to_float(m.group(1)) if m else None

def extract_retail_yoy(txt: str) -> Optional[float]:
    m = re.search(r"total retail sales of consumer goods .*? up by ([\\d\\.\\-]+)% year on year", txt, flags=re.I)
    return to_float(m.group(1)) if m else None

def extract_fai_yoy(txt: str) -> Optional[float]:
    m = re.search(r"investment in fixed assets .*? up by ([\\d\\.\\-]+)% year on year", txt, flags=re.I)
    return to_float(m.group(1)) if m else None

def extract_pmi_mfg(txt: str) -> Optional[float]:
    m = re.search(r"Manufacturing PMI was ([\\d\\.]+)%", txt, flags=re.I)
    return to_float(m.group(1)) if m else None

def extract_pmi_nonmfg(txt: str) -> Optional[float]:
    m = re.search(r"Non-manufacturing Business Activity Index was ([\\d\\.]+)%", txt, flags=re.I)
    return to_float(m.group(1)) if m else None

def extract_gdp_q_yoy(txt: str) -> Optional[float]:
    m = re.search(r"GDP .*? grew ([\\d\\.\\-]+) percent year on year", txt, flags=re.I)
    return to_float(m.group(1)) if m else None

def extract_unemp(txt: str) -> Optional[float]:
    m = re.search(r"urban surveyed unemployment rate was ([\\d\\.]+) percent", txt, flags=re.I)
    return to_float(m.group(1)) if m else None

def fetch_nbs_series(pattern: str, extract: str, start: dt.date, end: dt.date) -> List[Tuple[str, Optional[float]]]:
    items = fetch_nbs_rss()
    filtered = [(pd, t, l) for (pd, t, l) in items if pattern.lower() in t.lower()]
    s0 = iso(start - dt.timedelta(days=120))
    s1 = iso(end)
    out=[]
    for pub, title, link in filtered:
        if pub < s0 or pub > s1: 
            continue
        status, headers, body = fetch_url(link)
        if status != 200: 
            continue
        txt = strip_tags(body.decode("utf-8", errors="replace"))
        d = month_from_title(title) or quarter_from_title(title) or (pub[:7] + "-01")
        val=None
        if extract=="cpi_yoy": val=extract_cpi_yoy(txt)
        elif extract=="ppi_yoy": val=extract_ppi_yoy(txt)
        elif extract=="retail_yoy": val=extract_retail_yoy(txt)
        elif extract=="fai_yoy": val=extract_fai_yoy(txt)
        elif extract=="pmi_mfg": val=extract_pmi_mfg(txt)
        elif extract=="pmi_nonmfg": val=extract_pmi_nonmfg(txt)
        elif extract=="gdp_q_yoy": val=extract_gdp_q_yoy(txt)
        elif extract=="unemp_q1": val=extract_unemp(txt)
        out.append((d, val))
    # dedupe
    m={}
    for d,v in out: m[d]=v
    return sorted(m.items(), key=lambda x:x[0])

# ---------- Panel build ----------
def forward_fill_daily(raw: Dict[str, Dict[str, Optional[float]]], start: dt.date, end: dt.date) -> Dict[str, Dict[str, Optional[float]]]:
    daily = {k: {} for k in raw.keys()}
    for k, series in raw.items():
        dates_sorted = sorted(series.keys())
        last_v=None; j=0
        for day in daterange(start, end):
            ds = iso(day)
            while j < len(dates_sorted) and dates_sorted[j] <= ds:
                last_v = series[dates_sorted[j]]
                j += 1
            daily[k][ds] = last_v
    return daily

def derive_spread_bps(daily, out_key, a, b):
    if a not in daily or b not in daily: return
    daily[out_key] = {}
    for ds,av in daily[a].items():
        bv = daily[b].get(ds)
        daily[out_key][ds] = None if av is None or bv is None else (av-bv)*100.0

def derive_yoy_pct(daily, out_key, a, lag_days):
    if a not in daily: return
    daily[out_key] = {}
    for ds,av in daily[a].items():
        if av is None: 
            daily[out_key][ds]=None; continue
        lag_s = iso(parse_iso(ds) - dt.timedelta(days=lag_days))
        bv = daily[a].get(lag_s)
        daily[out_key][ds] = None if bv is None or bv==0 else (av/bv - 1.0)*100.0

def derive_diff_lag(daily, out_key, a, lag_days):
    if a not in daily: return
    daily[out_key] = {}
    for ds,av in daily[a].items():
        if av is None: 
            daily[out_key][ds]=None; continue
        lag_s = iso(parse_iso(ds) - dt.timedelta(days=lag_days))
        bv = daily[a].get(lag_s)
        daily[out_key][ds] = None if bv is None else (av-bv)

def write_wide_csv(path, dates, keys, daily):
    with open(path,"w",newline="",encoding="utf-8") as f:
        w=csv.writer(f); w.writerow(["date"]+keys)
        for ds in dates:
            w.writerow([ds] + ["" if daily.get(k,{}).get(ds) is None else daily[k][ds] for k in keys])

def write_wide_xlsx(path, dates, keys, daily, meta):
    if Workbook is None: return
    wb=Workbook()
    ws=wb.active; ws.title="daily_wide"
    ws.append(["date"]+keys)
    for ds in dates:
        ws.append([ds] + [daily.get(k,{}).get(ds) for k in keys])
    ws2=wb.create_sheet("meta")
    ws2.append(["key","name","provider","freq","unit"])
    for k in keys:
        m=meta.get(k,{})
        ws2.append([k, m.get("name",""), m.get("provider",""), m.get("freq",""), m.get("unit","")])
    wb.save(path)

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument("--registry", required=True)
    ap.add_argument("--out", default="out")
    ap.add_argument("--start", default="2026-04-01")
    ap.add_argument("--history_buffer_days", type=int, default=420)
    args=ap.parse_args()

    start=parse_iso(args.start)
    end=dt.date.today()
    fetch_start = start - dt.timedelta(days=args.history_buffer_days)

    with open(args.registry,"r",encoding="utf-8") as f:
        reg=json.load(f)

    out_dir=args.out
    raw_dir=os.path.join(out_dir,"raw")
    ensure_dir(out_dir); ensure_dir(raw_dir)

    series_defs=reg.get("series",[])
    derived_defs=reg.get("derived",[])

    raw_map={}
    meta={}

    start_year=fetch_start.year
    end_year=end.year

    for s in series_defs:
        key=s["key"]; provider=s.get("provider","")
        meta[key]={"name":s.get("name",""),"provider":provider,"freq":s.get("freq",""),"unit":s.get("unit","")}

        manual_path=os.path.join(raw_dir,f"{key}.csv")
        if provider=="manual":
            raw_map[key] = {d:v for d,v in read_manual_csv(manual_path)} if os.path.exists(manual_path) else {}
            continue

        try:
            rows=[]
            if provider=="stooq":
                p=s.get("params",{})
                rows=fetch_stooq(p["symbol"], fetch_start, end, p.get("interval","d"))
            elif provider=="hkma":
                p=s.get("params",{})
                rows=fetch_hkma(p["path"], p.get("field","efn_10y"), fetch_start, end, int(p.get("offset",0)))
            elif provider=="csd":
                p=s.get("params",{})
                rows=fetch_csd(str(p["tableId"]), fetch_start, end, p.get("lang","en"), int(p.get("full_series",1)), p.get("param",""))
            elif provider=="bls":
                p=s.get("params",{})
                rows=fetch_bls_v1(p["seriesId"], start_year, end_year)
            elif provider=="nbs_rss":
                p=s.get("params",{})
                rows=fetch_nbs_series(p.get("pattern",""), p.get("extract",""), fetch_start, end)
            write_csv(os.path.join(raw_dir,f"{key}.csv"), rows)
            raw_map[key] = {d:v for d,v in rows if d}
        except Exception as e:
            print(f"[WARN] {key} fetch failed: {e}", file=sys.stderr)
            raw_map[key] = {}

    daily = forward_fill_daily(raw_map, start, end)

    for d in derived_defs:
        op=d.get("op")
        if op=="spread_bps": derive_spread_bps(daily, d["key"], d["a"], d["b"])
        elif op=="yoy_pct": derive_yoy_pct(daily, d["key"], d["a"], int(d.get("lag_days",365)))
        elif op=="diff_lag": derive_diff_lag(daily, d["key"], d["a"], int(d.get("lag_days",30)))
        meta[d["key"]]={"name":d.get("name",""),"provider":"derived","freq":"daily","unit":d.get("unit","")}

    dates=[iso(d) for d in daterange(start,end)]
    keys=sorted(list(daily.keys()))

    wide_csv=os.path.join(out_dir,"daily_wide.csv")
    wide_xlsx=os.path.join(out_dir,"daily_wide.xlsx")
    write_wide_csv(wide_csv, dates, keys, daily)
    write_wide_xlsx(wide_xlsx, dates, keys, daily, meta)

    print(f"OK: wrote {wide_csv}")
    if Workbook is None:
        print("NOTE: openpyxl not available, skipped xlsx. Install: pip3 install openpyxl")
    else:
        print(f"OK: wrote {wide_xlsx}")

if __name__=="__main__":
    main()

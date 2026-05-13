#!/usr/bin/env python3
# macro_doctor.py
# Quick connectivity & parsing test for the keyless collector sources.
#
# Usage:
#   python3 macro_doctor.py
#
import json, datetime as dt, urllib.request, ssl, gzip, time

UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36"

def fetch(url, method="GET", data=None, headers=None, timeout=25):
    h = {"User-Agent": UA, "Accept":"*/*"}
    if headers: h.update(headers)
    req = urllib.request.Request(url, data=data, headers=h, method=method)
    ctx = ssl.create_default_context()
    with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
        status = getattr(resp, "status", 200)
        hdrs = dict(resp.headers.items())
        body = resp.read()
    if hdrs.get("Content-Encoding","").lower()=="gzip":
        try: body = gzip.decompress(body)
        except Exception: pass
    return status, hdrs, body

tests = [
  ("HKMA EFBN yield daily", "https://api.hkma.gov.hk/public/market-data-and-statistics/monthly-statistical-bulletin/efbn/efbn-yield-daily?offset=0", "json"),
  ("C&SD sample table (unemployment, id=6)", "https://www.censtatd.gov.hk/api/get.php?id=6&lang=en&full_series=1", "json"),
  ("NBS PressRelease RSS", "https://www.stats.gov.cn/english/PressRelease/rss.xml", "xml"),
  ("Stooq SPX CSV", "https://stooq.com/q/d/l/?s=%5Espx&d1=20260401&d2=20260513&i=d", "csv"),
  ("BLS v1 UNRATE", "https://api.bls.gov/publicAPI/v1/timeseries/data/", "bls_post"),
  ("ChinaMoney LPR (HTML)", "https://www.chinamoney.com.cn/english/bmklpr/", "html"),
  ("SAFE Reserves latest (HTML)", "https://www.safe.gov.cn/en/ForeignExchangeReserves/index.html", "html"),
]

def main():
    print("Macro Doctor - connectivity & basic parsing checks\\n")
    for name, url, typ in tests:
        try:
            if typ == "bls_post":
                payload = json.dumps({"seriesid":["LNS14000000"],"startyear":"2025","endyear":"2026"}).encode("utf-8")
                st, hdr, body = fetch(url, method="POST", data=payload, headers={"Content-Type":"application/json"})
            else:
                st, hdr, body = fetch(url)

            print(f"[OK] {name}: HTTP {st}, bytes={len(body)}")
            if typ == "json":
                import json as _j
                obj = _j.loads(body.decode("utf-8", errors="replace"))
                # show a hint of keys
                print("     keys:", list(obj.keys())[:6])
            elif typ == "xml":
                txt = body.decode("utf-8", errors="replace")
                print("     contains <item>:", "<item>" in txt)
            elif typ == "csv":
                txt = body.decode("utf-8", errors="replace").splitlines()
                print("     head:", txt[0] if txt else "(empty)")
                print("     rows:", max(0, len(txt)-1))
            elif typ == "html":
                txt = body.decode("utf-8", errors="replace")
                print("     snippet:", txt[:120].replace("\\n"," ") + "...")
            elif typ == "bls_post":
                txt = body.decode("utf-8", errors="replace")
                obj = json.loads(txt)
                s = obj.get("Results", {}).get("series", [{}])[0]
                print("     series:", s.get("seriesID"))
                print("     points:", len(s.get("data", [])))
            print()
        except Exception as e:
            print(f"[FAIL] {name}: {e}\\n")

if __name__ == "__main__":
    main()

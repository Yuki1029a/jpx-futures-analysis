"""test01.xlsx の先物ティック（レガシー形式 YYYYMMDD_D/_E）を読み、1 分足にして parquet に保存。"""
import re, time, openpyxl, pandas as pd, numpy as np
t0 = time.time(); wb = openpyxl.load_workbook("H:/マイドライブ/bank/test01.xlsx", read_only=True)
frames = []
for n in sorted(wb.sheetnames):
    m = re.match(r"^(\d{8})_([DE])$", n)
    if not m: continue
    ws = wb[n]; rows = list(ws.iter_rows(values_only=True))
    if len(rows) <= 1: print(n, "empty"); continue
    hdr = rows[0]; df = pd.DataFrame(rows[1:], columns=hdr)
    df.columns = ["hhmm", "price", "qty", "hhmmssms", "cumvol", "dir", "cluster"][:len(df.columns)]
    df["date"] = m.group(1); df["session"] = m.group(2); frames.append(df)
tk = pd.concat(frames, ignore_index=True); print("ticks", len(tk), f"({time.time()-t0:.0f}s)", tk.date.min(), tk.date.max(), tk.groupby("session").size().to_dict())
tk["price"] = pd.to_numeric(tk.price, errors="coerce"); tk["qty"] = pd.to_numeric(tk.qty, errors="coerce"); tk = tk.dropna(subset=["price", "qty"])
# 時刻: ナイトは 17:00→翌 06:00。取引日 = セッション開始日。時刻を分に（ナイトの 0:00 以降は +24h）
hm = tk.hhmm.astype(str).str.extract(r"(\d{1,2}):(\d{2})").astype(float); tk["min"] = hm[0] * 60 + hm[1]
tk.loc[(tk.session == "E") & (tk["min"] < 12 * 60), "min"] += 24 * 60
tk["sess_id"] = tk.date + "_" + tk.session
bar = tk.groupby(["sess_id", "date", "session", "min"]).agg(o=("price", "first"), h=("price", "max"), l=("price", "min"), c=("price", "last"), v=("qty", "sum"), n=("price", "size")).reset_index()
bar.to_parquet("C:/Users/kawai/AppData/Local/Temp/tick/bars_1min.parquet"); print("bars", len(bar), "sessions", bar.sess_id.nunique())
print(bar.groupby("session").agg(sessions=("sess_id", "nunique"), min_start=("min", "min"), min_end=("min", "max"), vol=("v", "sum")).to_string())
print(tk.head(3).to_string())

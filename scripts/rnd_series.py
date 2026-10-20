"""11月限のリスク中立分布を日次（引け近傍）で計算し、系列を JSON に保存する。

使い方: python scripts/rnd_series.py <出力json> [cm=2611]
各営業日は 15:15 以前の最終スナップショット。最終日が土日・夜間明けなら、その直前の夜間立会の最終スナップショット（06:15 以前）も「夜間引け」として追加する。
"""
import io
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent)); sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import pandas as pd

from data import iv_views
from rnd_lib import NIGHT_CLOSE, analyze_key, day_close_snap, snap_for

out_path = Path(sys.argv[1]); cm = sys.argv[2] if len(sys.argv) > 2 else "2611"
start = sys.argv[3] if len(sys.argv) > 3 else "20261009"
days = [d for d in sorted(iv_views.list_available_days(60)) if d >= start]
rows = []
for d in days:
    wd = pd.Timestamp(d).weekday()
    if wd < 5:
        k = day_close_snap(d, cm)  # 気配時刻 15:30:59 以前の最終（夜間の気配は使わない）
        if k:
            r = analyze_key(k, cm); r["label"] = f"{d[4:6]}/{d[6:8]} 引け"; rows.append(r); print(r["label"], r["ts"], round(r["F"]), round(r["atm"] * 100, 2), flush=True)
# 最終営業日の夜間引け（翌暦日の 06:00:59 以前、気配時刻基準）は参考値として末尾に付ける
last = days[-1]
k = snap_for(last, NIGHT_CLOSE, cm)
if k and (pd.Timestamp(last).weekday() >= 5 or last > rows[-1]["ts"][:10].replace("-", "")):
    r = analyze_key(k, cm); r["label"] = f"{last[4:6]}/{last[6:8]} 夜間引け（参考）"; rows.append(r); print(r["label"], r["ts"], round(r["F"]), round(r["atm"] * 100, 2))
json.dump(rows, open(out_path, "w"))
print("saved", out_path, len(rows))

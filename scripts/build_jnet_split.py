"""JPX 日報（siop_dyr PDF）から行使価格別の立会内／立会外取引高を作り、R2 とローカルに保存する。

保存: R2 jnet/siop/YYYYMMDD.parquet、ローカル cache/jnet/siop/YYYYMMDD.parquet（列は data/jnet_split.py 参照）
使い方: python scripts/build_jnet_split.py [--day YYYYMMDD | --since YYYYMMDD] [--upload] [--force] [--workers N]
既定: 直近 7 暦日の平日のうち未作成の日。日報は翌営業日 09:00 頃公表のため、未公表の日はスキップ（次回の実行で拾う）。
新形式 zip（2026-01-23〜）が無い日は旧形式の統合 PDF を使う。
"""
from __future__ import annotations

import argparse
import io
import logging
import sys
import zipfile
from concurrent.futures import ProcessPoolExecutor
from datetime import date, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT)); sys.path.insert(0, str(ROOT / "scripts"))
logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s"); log = logging.getLogger("jnet")


def siop_pdf(d: date) -> bytes | None:
    from backfill_daily_oi_from_pdf import _split_old_pdf, get_old_pdf, get_zip
    ds = d.strftime("%Y%m%d"); z = get_zip(d)
    if z is not None:
        with zipfile.ZipFile(io.BytesIO(z)) as zf:
            names = {n.rsplit("/", 1)[-1]: n for n in zf.namelist()}
            n = names.get(f"siop_dyr_{ds}.pdf")
            return zf.read(n) if n else None
    old = get_old_pdf(d)
    return _split_old_pdf(old)[1] if old is not None else None


def build_one(ds: str, upload: bool) -> str:
    from data.jnet_split import LOCAL, R2_PREFIX, build_from_siop
    from data.r2_storage import r2_put
    d = date(int(ds[:4]), int(ds[4:6]), int(ds[6:8]))
    pdf = siop_pdf(d)
    if pdf is None:
        return f"{ds}: 日報なし"
    rd, df = build_from_siop(pdf)
    if rd != d:
        return f"{ds}: 日付不一致 {rd}"
    LOCAL.mkdir(parents=True, exist_ok=True); buf = io.BytesIO(); df.to_parquet(buf, index=False); body = buf.getvalue()
    (LOCAL / f"{ds}.parquet").write_bytes(body)
    if upload:
        r2_put(f"{R2_PREFIX}/{ds}.parquet", body)
    op = df[df["product"] == "op225"]; mi = df[df["product"] == "mini"]
    return f"{ds}: 225OP 立会内 {int(op.auction_vol.sum()):,} 立会外 {int(op.jnet_vol.sum()):,} / ミニ 立会内 {int(mi.auction_vol.sum()):,} 立会外 {int(mi.jnet_vol.sum()):,}"


def exists(ds: str, upload: bool) -> bool:
    from data.jnet_split import LOCAL, R2_PREFIX
    if (LOCAL / f"{ds}.parquet").exists():
        return True
    if upload:
        from data.r2_storage import r2_get
        return r2_get(f"{R2_PREFIX}/{ds}.parquet") is not None
    return False


def main() -> int:
    ap = argparse.ArgumentParser(); ap.add_argument("--day"); ap.add_argument("--since"); ap.add_argument("--upload", action="store_true")
    ap.add_argument("--force", action="store_true"); ap.add_argument("--workers", type=int, default=1)
    a = ap.parse_args(); today = date.today()
    if a.day:
        days = [a.day]
    elif a.since:
        d0 = date(int(a.since[:4]), int(a.since[4:6]), int(a.since[6:8]))
        days = [(d0 + timedelta(days=i)).strftime("%Y%m%d") for i in range((today - d0).days + 1) if (d0 + timedelta(days=i)).weekday() < 5]
    else:
        days = [(today - timedelta(days=b)).strftime("%Y%m%d") for b in range(7, -1, -1) if (today - timedelta(days=b)).weekday() < 5]
    todo = [ds for ds in days if a.force or not exists(ds, a.upload)]
    log.info("対象 %d 日（未作成 %d 日）", len(days), len(todo))
    if a.workers > 1 and len(todo) > 1:
        with ProcessPoolExecutor(a.workers) as ex:
            for msg in ex.map(build_one, todo, [a.upload] * len(todo)):
                log.info(msg)
    else:
        for ds in todo:
            try:
                log.info(build_one(ds, a.upload))
            except Exception as e:  # noqa: BLE001
                log.warning("%s: failed: %s", ds, e)
    return 0


if __name__ == "__main__":
    sys.exit(main())

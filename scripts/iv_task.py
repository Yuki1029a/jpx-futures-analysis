"""Windows タスクスケジューラから15分おきに呼ばれ、取引時間内だけ QRI IV を収集して R2 へ保存する。

GitHub Actions の定期実行は間引かれて実質2〜3時間おきになるため、この PC から補完する。
取引時間（JST）: 日中 08:45-15:30、夜間 16:30-翌06:15（QRI の15分遅延配信ぶん余裕を持たせる）。
ログ: cache/iv_task.log
"""
import subprocess
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def in_session(now: datetime) -> bool:
    m, dow = now.hour * 60 + now.minute, now.weekday()  # 0=月 … 6=日
    # 日中立会 08:45-15:45（配信が15分遅延のため、引け気配を取るには 16:15 まで取得を続ける）
    # 夜間立会 17:00-翌06:00（同じく 06:15 まで）。16:15-17:00 は板が消えるため取得しない
    day = dow <= 4 and 8 * 60 + 45 <= m <= 16 * 60 + 15
    eve = dow <= 4 and m >= 17 * 60
    morning = 1 <= dow <= 5 and m <= 6 * 60 + 15
    return day or eve or morning


if __name__ == "__main__":
    now = datetime.now()
    if not in_session(now) and "--force" not in sys.argv:
        sys.exit(0)
    log = ROOT / "cache" / "iv_task.log"
    log.parent.mkdir(exist_ok=True)
    with open(log, "a", encoding="utf-8") as f:
        f.write(f"--- {now:%Y-%m-%d %H:%M:%S}\n")
        f.flush()
        r = subprocess.run([sys.executable, str(ROOT / "scripts" / "fetch_qri_iv.py"), "--r2", "--no-local"],
                           stdout=f, stderr=subprocess.STDOUT, timeout=600)
        f.write(f"exit {r.returncode}\n")
        # 平日 06:00-06:30 の回（夜間引け後）に J-Quants の前営業日分（立会内／立会外の内訳など）を取得して R2 へ
        if now.weekday() <= 4 and 6 * 60 <= now.hour * 60 + now.minute < 6 * 60 + 30:
            f.write("--- jquants options\n"); f.flush()
            r2 = subprocess.run([sys.executable, str(ROOT / "scripts" / "fetch_jquants_options.py"), "--upload"], stdout=f, stderr=subprocess.STDOUT, timeout=600)
            f.write(f"exit {r2.returncode}\n")

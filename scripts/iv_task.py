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
    day = dow <= 4 and 8 * 60 + 45 <= m <= 15 * 60 + 30
    eve = dow <= 4 and m >= 16 * 60 + 30
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

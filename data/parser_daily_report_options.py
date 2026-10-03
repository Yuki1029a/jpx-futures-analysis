"""JPX 日々の統計 (Daily_Report_OSE_{yyyymmdd}.zip) 内 siop_dyr_{yyyymmdd}.pdf
(指数オプション日報) から日経225オプションの銘柄別データを抽出する。

建玉残高表 xlsx が取得できなかった日の代替ソースとして使用する。

PDF 構造 (2026-09-29 実測, A3 横 1191x842pt, 112 ページ):
  各ページ左上ヘッダ (x0≈92):
    行1  商品名        例: 日経225オプション / 日経225ミニオプション / TOPIXオプション /
                       JPX日経インデックス400オプション / 東証銀行業株価指数オプション /
                       東証REIT指数オプション
    行2  市場区分      競争売買市場 (AuctionMarket) | J-NET市場 (J-NETMarket)
    行3  銘柄区分      取引成立銘柄 (TradeExecutedIssues)   ※9/29-10/1 は全ページこの区分のみ
    右上 日付          "2026年" "9月29日(火曜日)"
  ページ順: [競争売買市場] を全商品分 → [J-NET市場] を全商品分。
    9/29: p1-61 日経225OP 競争 (p1 プット開始, p36 途中からコール),
          p62-82 ミニOP 競争, p83-95 TOPIX OP 競争, p96-98 他指数 OP 競争,
          p99-106 日経225OP J-NET (p99 プット, p103 途中からコール), p107-112 他 J-NET。
  セクション: 本文中の "プットオプション PutOptions" / "コールオプション CallOptions" 行。
    同一商品・同一市場内でページをまたいで継続する (継続ページにはマーカーなし)。
  データ行 (1 銘柄 = 1 行, 行の折返しなし):
    競争売買市場: 限月(yyyymm) 最終日(mm.dd) 権利行使価格 コード
                  夜間 始値 高値 安値 終値 | 日中 始値 高値 安値 終値 | 前日比較
                  取引高(概算) 取引金額(円,概算) 清算価格 権利行使数量 建玉残高(概算)
      ・前日比較は符号と数値が別トークン ("- 1.0000") になる。
      ・未取引で建玉のみの銘柄も掲載 (取引高 "…")。
    J-NET市場:    限月 最終日 権利行使価格 コード 始値 高値 安値 終値 取引高 取引金額
      (清算価格・権利行使数量・建玉残高なし。取引成立銘柄のみ)
    欠損値は "…"。ミニオプションは限月列が yyyymmdd (週次) になる。
  列判定: ヘッダ語 (権利行使[価格] / 取引高 / 取引金額 / 清算価格 / 権利行使[数量] /
    建玉残高 等) の右端 x1 に対し、各トークンの右端 x1 が最も近い列へ割り当てる
    (数値は右寄せのため右端基準が桁数に依存せず安定)。

検証 (9/29):
  競争売買市場の建玉残高 = 建玉残高表 別紙1 当日建玉残高、
  競争売買 + J-NET の取引高 = 建玉残高表 別紙1 取引高 (別紙1 は直近3限月のみ)。
"""
from __future__ import annotations

import io
import logging
import re
from datetime import date

import pdfplumber

logger = logging.getLogger(__name__)

_TARGET_PRODUCT = "日経225オプション"
_MARKETS = {"競争売買市場": "auction", "J-NET市場": "jnet"}
_SECTIONS = {"プットオプション": "PUT", "コールオプション": "CALL"}
_KNOWN_KIND = "取引成立銘柄"

_DATE_RE = re.compile(r"(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日")
_CONTRACT_RE = re.compile(r"^\d{6}$")
_LASTDAY_RE = re.compile(r"^\d{2}\.\d{2}$")

# ヘッダ語 → 列名。code/close/chg は誤割当防止用のダミー列
_HEADER_COLS = {
    "コード": "code", "終値": "close", "前日比較": "chg",
    "取引高": "volume", "取引金額": "value", "清算価格": "settle",
    "建玉残高": "oi",
}
_HEADER_X_LIMIT = 120.0   # 商品名・市場区分・銘柄区分は x0 がこの左側
_ROW_TOL = 3.0            # 同一行とみなす top の許容差 (pt)


def parse_siop_pdf(content: bytes) -> tuple[date, list[dict]]:
    """siop_dyr PDF から日経225オプション (ラージ) の銘柄別行を返す。

    Returns:
      (report_date, rows)。rows の各要素:
        contract_month 'YYMM', option_type 'PUT'|'CALL', strike_price int,
        trading_volume int, trading_value_yen int, open_interest int,
        settlement_price float|None, exercised int, market 'auction'|'jnet'
      '…' は 0 (settlement_price は None)。J-NET 行は OI/清算/権利行使なし。
    ミニオプション・TOPIX 等の他商品ページは除外する。
    """
    report_date: date | None = None
    rows: list[dict] = []
    section: str | None = None
    prev_key: tuple[str, str] | None = None

    with pdfplumber.open(io.BytesIO(content)) as pdf:
        for pno, page in enumerate(pdf.pages, start=1):
            words = page.extract_words(x_tolerance=1.5)
            lines = _group_lines(words)
            product, market_jp, kind = _page_header(lines)
            if product != _TARGET_PRODUCT or market_jp not in _MARKETS:
                continue
            if report_date is None:
                report_date = _find_date(words)
            if kind != _KNOWN_KIND:
                logger.warning("siop p%d: unexpected issue kind %r", pno, kind)
            market = _MARKETS[market_jp]
            key = (product, market_jp)
            if key != prev_key:
                section = None          # 商品/市場が変わればセクション状態をリセット
                prev_key = key

            anchors = _header_anchors(lines)
            if "volume" not in anchors or "strike" not in anchors:
                logger.warning("siop p%d: header anchors not found", pno)
                continue

            for line in lines:
                first = line[0]["text"]
                if first in _SECTIONS:
                    section = _SECTIONS[first]
                    continue
                if not (_CONTRACT_RE.match(first) and len(line) >= 3
                        and _LASTDAY_RE.match(line[1]["text"])):
                    continue
                if section is None:
                    logger.warning("siop p%d: data row before section marker", pno)
                    continue
                cols = _assign_columns(line[2:], anchors)
                rows.append({
                    "contract_month": first[2:],
                    "option_type": section,
                    "strike_price": _to_int(cols.get("strike")),
                    "trading_volume": _to_int(cols.get("volume")),
                    "trading_value_yen": _to_int(cols.get("value")),
                    "open_interest": _to_int(cols.get("oi")),
                    "settlement_price": _to_float(cols.get("settle")),
                    "exercised": _to_int(cols.get("exercised")),
                    "market": market,
                })

    if report_date is None:
        raise ValueError("siop PDF: 日経225オプションページ/日付が見つかりません")
    return report_date, rows


# ---------------------------------------------------------------- helpers

def _group_lines(words: list[dict]) -> list[list[dict]]:
    """top が近い語を同一行にまとめ、行内は x0 順に並べる。"""
    lines: list[list[dict]] = []
    for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
        if lines and abs(lines[-1][0]["top"] - w["top"]) <= _ROW_TOL:
            lines[-1].append(w)
        else:
            lines.append([w])
    for ln in lines:
        ln.sort(key=lambda w: w["x0"])
    return lines


def _page_header(lines: list[list[dict]]) -> tuple[str | None, str | None, str | None]:
    """左上ヘッダから (商品名, 市場区分, 銘柄区分) を返す。"""
    product = market = kind = None
    for ln in lines:
        w = ln[0]
        if w["x0"] > _HEADER_X_LIMIT:
            continue
        t = w["text"]
        if t.endswith("オプション") and product is None:
            product = t
        elif t in _MARKETS:
            market = t
        elif t.endswith("銘柄"):
            kind = t
        if product and market and kind:
            break
    return product, market, kind


def _find_date(words: list[dict]) -> date | None:
    text = " ".join(w["text"] for w in words if w["top"] < 200)
    m = _DATE_RE.search(text)
    return date(int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else None


def _header_anchors(lines: list[list[dict]]) -> dict[str, float]:
    """ヘッダ語の右端 x1 を列名ごとに返す。

    権利行使 は2箇所 (左=価格 → strike, 右=数量 → exercised) あり x0 で区別する。
    終値 は最右 (日中終値) を採用する。
    """
    anchors: dict[str, float] = {}
    for ln in lines:
        for w in ln:
            t = w["text"]
            if t == "権利行使":
                anchors["strike" if w["x0"] < 400 else "exercised"] = w["x1"]
            elif t in _HEADER_COLS:
                name = _HEADER_COLS[t]
                if name == "close":
                    anchors[name] = max(anchors.get(name, 0.0), w["x1"])
                else:
                    anchors.setdefault(name, w["x1"])
        if ln[0]["text"] in _SECTIONS:
            break   # ヘッダ領域終了
    return anchors


def _assign_columns(words: list[dict], anchors: dict[str, float]) -> dict[str, str]:
    """各トークンを右端 x1 が最も近いヘッダ列へ割り当てる (列ごとに最初の語)。"""
    out: dict[str, str] = {}
    for w in words:
        name = min(anchors, key=lambda k: abs(anchors[k] - w["x1"]))
        out.setdefault(name, w["text"])
    return out


def _to_int(s: str | None) -> int:
    if not s or s == "…":
        return 0
    try:
        return int(s.replace(",", ""))
    except ValueError:
        return 0


def _to_float(s: str | None) -> float | None:
    if not s or s == "…":
        return None
    try:
        return float(s.replace(",", ""))
    except ValueError:
        return None


if __name__ == "__main__":
    import sys
    import zipfile
    from collections import Counter

    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    if len(sys.argv) < 2:
        print("usage: python parser_daily_report_options.py "
              "<Daily_Report_OSE_yyyymmdd.zip | siop_dyr_yyyymmdd.pdf>")
        sys.exit(1)
    path = sys.argv[1]
    if path.lower().endswith(".zip"):
        with zipfile.ZipFile(path) as zf:
            name = next(n for n in zf.namelist()
                        if n.startswith("siop_dyr_") and "flex" not in n)
            content = zf.read(name)
    else:
        with open(path, "rb") as f:
            content = f.read()
    d, rows = parse_siop_pdf(content)
    print(f"report_date={d} rows={len(rows)}")
    cnt = Counter((r["market"], r["contract_month"], r["option_type"]) for r in rows)
    for k in sorted(cnt):
        print(f"  {k[0]:7s} {k[1]} {k[2]:4s} {cnt[k]:5d}")
    for mk in ("auction", "jnet"):
        for ot in ("PUT", "CALL"):
            sub = [r for r in rows if r["market"] == mk and r["option_type"] == ot]
            print(f"  {mk:7s} {ot:4s} volume={sum(r['trading_volume'] for r in sub):,} "
                  f"value_yen={sum(r['trading_value_yen'] for r in sub):,} "
                  f"oi={sum(r['open_interest'] for r in sub):,}")

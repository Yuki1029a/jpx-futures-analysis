"""JPX 日々の統計「指数先物 日報」PDF (sif_dyr_{yyyymmdd}.pdf) パーサ.

建玉残高表 Excel が欠落した日の代替データソース。1商品1ページ構成で、
前半ページが競争売買市場(Auction)、後半ページが J-NET市場 取引成立銘柄。

列境界は PDF の縦罫線 (page.lines) から取得し、見出し語 (取引高 / 取引金額 /
清算数値 / 建玉残高) で列を同定する。罫線が取れない場合は見出し語の中心から
境界を推定する。前日比較の符号 ('+', '-') と数値が別トークンに分かれても、
x 座標で列に割り付けるため列ずれは起きない。

Excel (建玉残高表) の 取引高 は 競争売買市場 + J-NET市場 の合算値。
本モジュールの trading_volume / trading_value_mil も同じ定義で合算する。
内訳は auction_* / jnet_* に保持する。
"""
from __future__ import annotations

import io
import re
import sys
import zipfile
from datetime import date
from typing import Any

import pdfplumber

# 建玉残高表 (parser_daily_oi.py) と同じ商品コード。他商品は None。
PRODUCT_CODES: dict[str, str] = {
    "日経225先物": "NK225F",
    "日経225mini": "NK225MF",
    "日経225マイクロ先物": "NK225MicroF",
    "TOPIX先物": "TOPIXF",
    "ミニTOPIX先物": "MiniTOPIXF",
    "JPX日経インデックス400先物": "JPX400F",
}

_DATE_RE = re.compile(r"(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日")
_YYYYMM_RE = re.compile(r"^\d{6}$")
_MISSING = "…"
_HEADER_KEYS = ("取引高", "取引金額", "清算数値", "建玉残高", "うちストラテジー", "前日比較")


def _num(s: str | None) -> float | None:
    """'65,830.00' -> 65830.0。'…' / 空 -> None。"""
    if not s:
        return None
    s = s.replace(",", "").replace(" ", "")
    if s in ("", _MISSING, "-", "+"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _int0(s: str | None) -> int:
    """数値文字列 -> int。'…' は 0。"""
    v = _num(s)
    return int(round(v)) if v is not None else 0


def _group_rows(words: list[dict], tol: float = 3.0) -> list[list[dict]]:
    """words を top 座標でまとめて行にする (各行 x0 昇順)。"""
    rows: list[list[dict]] = []
    for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
        if rows and abs(rows[-1][0]["top"] - w["top"]) <= tol:
            rows[-1].append(w)
        else:
            rows.append([w])
    return [sorted(r, key=lambda w: w["x0"]) for r in rows]


def _column_bounds(page, header_words: list[dict]) -> list[float]:
    """列境界 x 座標 (昇順)。縦罫線優先、無ければ見出し語中心の中点で代替。"""
    xs = sorted({round(l["x0"], 1) for l in page.lines
                 if abs(l["x0"] - l["x1"]) < 1.0 and (l["bottom"] - l["top"]) > 100})
    if len(xs) >= 6:
        return xs
    centers = sorted({(w["x0"] + w["x1"]) / 2 for w in header_words if w["text"] in _HEADER_KEYS})
    return [0.0] + [(a + b) / 2 for a, b in zip(centers, centers[1:])] + [float(page.width)]


def _col_index(bounds: list[float], x: float) -> int | None:
    """x が属する列 index。境界外は None。"""
    for i in range(len(bounds) - 1):
        if bounds[i] <= x < bounds[i + 1]:
            return i
    return None


def _label_columns(bounds: list[float], header_words: list[dict]) -> dict[str, int]:
    """見出し語を列に割り付け、論理名 (volume/value/settlement/oi) -> 列index を返す。

    「うちストラテジー」列には英語見出し '取引高' も重なるため除外する。
    """
    labels: dict[int, str] = {}
    for w in header_words:
        ci = _col_index(bounds, (w["x0"] + w["x1"]) / 2)
        if ci is not None:
            labels[ci] = labels.get(ci, "") + w["text"]
    out: dict[str, int] = {}
    for ci in sorted(labels):
        lab = labels[ci]
        if "ストラテジー" in lab:
            continue
        if "取引高" in lab and "volume" not in out:
            out["volume"] = ci
        elif "取引金額" in lab and "value" not in out:
            out["value"] = ci
        elif "清算数値" in lab:
            out["settlement"] = ci
        elif "建玉残高" in lab:
            out["oi"] = ci
    return out


def _page_title(words: list[dict]) -> str:
    """左上の商品名 (空白除去)。右上の注記 (※…) は除く。"""
    ws = [w for w in words if w["top"] < 100 and w["x0"] < 320 and not w["text"].startswith("※")]
    return "".join(w["text"] for w in ws)


def _parse_page(page) -> tuple[str, str, date | None, list[dict[str, Any]]]:
    """1ページ解析。(product_jp, session 'auction'|'jnet', report_date, rows) を返す。"""
    words = page.extract_words(x_tolerance=1.5, keep_blank_chars=False)
    text = page.extract_text() or ""
    session = "jnet" if "J-NET市場" in text else "auction"
    m = _DATE_RE.search(text)
    rdate = date(int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else None
    title = _page_title(words)

    rows = _group_rows(words)
    data_rows = [r for r in rows if _YYYYMM_RE.match(r[0]["text"]) and r[0]["x0"] < 120]  # J-NET頁は x0≈80
    if not data_rows:
        return title, session, rdate, []
    first_top = data_rows[0][0]["top"]
    header_words = [w for w in words if 140 < w["top"] < first_top - 2]
    bounds = _column_bounds(page, header_words)
    cols = _label_columns(bounds, header_words)
    if "volume" not in cols:
        raise ValueError(f"取引高列を同定できません: title={title}")

    out: list[dict[str, Any]] = []
    for r in data_rows:
        cells: dict[int, str] = {}
        for w in r:
            ci = _col_index(bounds, (w["x0"] + w["x1"]) / 2)
            if ci is not None:
                cells[ci] = cells.get(ci, "") + w["text"]
        ltd = r[1]["text"] if len(r) > 1 and re.fullmatch(r"\d{2}\.\d{2}", r[1]["text"]) else None
        code = next((w["text"] for w in r if re.fullmatch(r"\d{9}", w["text"])), None)
        out.append({
            "contract_month": r[0]["text"][2:],
            "last_trading_day": ltd,
            "code": code,
            "volume": _int0(cells.get(cols["volume"])),
            "value_mil": _int0(cells.get(cols["value"])) if "value" in cols else 0,
            "settlement": _num(cells.get(cols["settlement"])) if "settlement" in cols else None,
            "oi": _int0(cells.get(cols["oi"])) if "oi" in cols else 0,
        })
    return title, session, rdate, out


def _new_row(title: str, pr: dict[str, Any]) -> dict[str, Any]:
    return {
        "product_jp": title,
        "product": PRODUCT_CODES.get(title),
        "contract_month": pr["contract_month"],
        "trading_volume": 0,
        "trading_value_mil": 0,
        "auction_volume": 0,
        "jnet_volume": 0,
        "auction_value_mil": 0,
        "jnet_value_mil": 0,
        "settlement_price": None,
        "open_interest": 0,
        "last_trading_day": pr["last_trading_day"],
        "code": pr["code"],
    }


def parse_sif_pdf(content: bytes) -> tuple[date, list[dict[str, Any]]]:
    """sif_dyr PDF 全ページを解析し (report_date, rows) を返す。

    rows の各要素:
      product_jp, product (PRODUCT_CODES の値 or None), contract_month 'YYMM',
      trading_volume (= auction + J-NET), trading_value_mil (= auction + J-NET, 百万円),
      auction_volume, jnet_volume, auction_value_mil, jnet_value_mil,
      settlement_price (清算数値, '…' は None), open_interest (建玉残高, '…' は 0),
      last_trading_day 'mm.dd', code (9桁銘柄コード).
    """
    merged: dict[tuple[str, str], dict[str, Any]] = {}
    report_date: date | None = None
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        for page in pdf.pages:
            title, session, rdate, prows = _parse_page(page)
            report_date = report_date or rdate
            for pr in prows:
                key = (title, pr["contract_month"])
                row = merged.setdefault(key, _new_row(title, pr))
                if session == "auction":
                    row["auction_volume"] += pr["volume"]
                    row["auction_value_mil"] += pr["value_mil"]
                    row["settlement_price"] = pr["settlement"]
                    row["open_interest"] = pr["oi"]
                else:
                    row["jnet_volume"] += pr["volume"]
                    row["jnet_value_mil"] += pr["value_mil"]
                row["trading_volume"] = row["auction_volume"] + row["jnet_volume"]
                row["trading_value_mil"] = row["auction_value_mil"] + row["jnet_value_mil"]
    if report_date is None:
        raise ValueError("報告日をPDFヘッダから取得できません")
    return report_date, list(merged.values())


def extract_sif_from_zip(content: bytes) -> bytes:
    """Daily_Report_OSE_{d}.zip から sif_dyr_{d}.pdf (flex を除く) を取り出す。"""
    with zipfile.ZipFile(io.BytesIO(content)) as z:
        names = [n for n in z.namelist()
                 if n.startswith("sif_dyr_") and "flex" not in n and n.endswith(".pdf")]
        if not names:
            raise FileNotFoundError("sif_dyr_*.pdf が zip 内にありません")
        return z.read(names[0])


def load_sif_rows(path: str) -> tuple[date, list[dict[str, Any]]]:
    """zip または pdf のパスから parse_sif_pdf を実行する。"""
    with open(path, "rb") as f:
        content = f.read()
    if path.lower().endswith(".zip"):
        content = extract_sif_from_zip(content)
    return parse_sif_pdf(content)


if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    if len(sys.argv) < 2:
        print("usage: python data/parser_daily_report_futures.py <zip_or_pdf_path>")
        sys.exit(1)
    d, rows = load_sif_rows(sys.argv[1])
    print(f"report_date={d} rows={len(rows)}")
    print("product\tproduct_jp\tcm\tvolume(auction+jnet)\tauction\tjnet\tvalue_mil\tsettle\toi")
    for r in rows:
        print(f"{r['product'] or '-'}\t{r['product_jp']}\t{r['contract_month']}\t{r['trading_volume']}\t"
              f"{r['auction_volume']}\t{r['jnet_volume']}\t{r['trading_value_mil']}\t"
              f"{r['settlement_price']}\t{r['open_interest']}")

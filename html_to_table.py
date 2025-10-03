# html_to_sheet.py
# URL固定 → スクレイピング → JSON化 → スプレッドシートへ追記
from typing import List, Dict
from urllib.parse import urljoin
from datetime import datetime
import re
import requests
from bs4 import BeautifulSoup, Tag

# === ハードコーディング（必要に応じて変更）========================
URL = "https://www.mhlw.go.jp/web/t_doc?dataId=81aa1263&dataType=0"
IFRAME_FIRST = True  # t_doc系でiframe内に本文があるため True 推奨

SERVICE_ACCOUNT_JSON = "./service_account.json"  # サービスアカウントjson
SPREADSHEET_ID = "1iHKZn9y-AZkeO9TUf0kN3_3GBIthwtTr4i19tNStg8U"  # 例: 1abcdEFG...
SHEET_TITLE = "更新情報一覧"  # 追記先シート名
# ===============================================================


# ---------------- 取得 ----------------
def fetch_html(url: str, iframe_first: bool = False) -> str:
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    r.encoding = r.apparent_encoding
    outer_html = r.text
    if not iframe_first:
        return outer_html

    soup = BeautifulSoup(outer_html, "lxml")
    iframe = soup.find("iframe", src=True)
    if not iframe:
        return outer_html

    inner_url = urljoin(url, iframe["src"])
    r2 = requests.get(inner_url, timeout=30)
    r2.raise_for_status()
    r2.encoding = r2.apparent_encoding
    return r2.text


def pick_contents_node(html: str) -> Tag | None:
    soup = BeautifulSoup(html, "lxml")
    for sel in (
        "html > body.body > div.wrapper > div.main > div#contents",
        "html > body.body > div.wrapper > div.main > div.contents",  # 保険
        "#contents",
        ".contents",
    ):
        node = soup.select_one(sel)
        if node:
            return node
    return None


# -------------- 探索（table→行→セル） --------------
def collect_bon_tables(contents: Tag) -> List[Tag]:
    tables: List[Tag] = []
    candidates = contents.select(":scope > div[id]") or contents.select("div[id]")
    for block in candidates:
        for fr in block.select("div.table_frame"):
            wrappers = fr.select(
                "div.table_wrpper, div.table_wrapper, div.table-wrapper"
            )
            if wrappers:
                for wp in wrappers:
                    tables.extend(wp.select("table.b-on"))
            else:
                tables.extend(fr.select("table.b-on"))
    return tables


def classless_trs_of_table(tbl: Tag) -> List[Tag]:
    tbody = tbl.find("tbody") or tbl
    rows: List[Tag] = []
    for tr in tbody.find_all("tr", recursive=False):
        if not tr.get("class"):
            rows.append(tr)
    return rows


def td_p_texts(tr: Tag) -> List[str]:
    cells: List[str] = []
    for td in tr.find_all("td", recursive=False):
        texts = [
            " ".join(p.stripped_strings).replace("\u3000", " ").strip()
            for p in td.find_all("p")
        ]
        texts = [t for t in texts if t]
        cells.append(" ".join(texts))
    return cells


# ヘッダ行は除外
EXPECTED_HEADER = [
    "粘膜に使用されることがない化粧品のうち洗い流すもの",
    "粘膜に使用されることがない化粧品のうち洗い流さないもの",
    "粘膜に使用されることがある化粧品",
]


def _norm(s: str) -> str:
    s = (s or "").replace("\u3000", " ")
    return re.sub(r"\s+", " ", s).strip()


def is_header_row(cells: List[str]) -> bool:
    return [_norm(c) for c in cells] == [_norm(x) for x in EXPECTED_HEADER]


def build_tables_rows(url: str, iframe_first: bool = False) -> List[List[List[str]]]:
    html = fetch_html(url, iframe_first=iframe_first)
    contents = pick_contents_node(html)
    if contents is None:
        raise RuntimeError("contents（id=contents / class=contents）が見つかりません。")
    tables = collect_bon_tables(contents)
    out: List[List[List[str]]] = []
    for tbl in tables:
        rows_1table: List[List[str]] = []
        for tr in classless_trs_of_table(tbl):
            row = td_p_texts(tr)
            if is_header_row(row):
                continue
            rows_1table.append(row)
        out.append(rows_1table)
    return out


# -------------- 行 → JSON(dict) --------------
NUM_RE = re.compile(r"^\d+(?:\.\d+)?$")


def strip_units_and_note(val: str) -> tuple[str, str, str]:
    """値から 合計量として / 国際単位 / g(ｇ) を除去 → (clean, unit, note)"""
    s = _norm(val or "")
    note = ""
    if "合計量として" in s:
        s = s.replace("合計量として", "").strip()
        note = "合計量として"
    unit = ""
    if "国際単位" in s:
        unit = "国際単位"
        s = s.replace("国際単位", "")
    if re.search(r"[gｇ]", s):
        if not unit:
            unit = "g"
        s = re.sub(r"[gｇ]", "", s)
    s = _norm(s)
    if NUM_RE.fullmatch(s) and not unit:
        unit = "g"
    return s, unit, note


def strip_units_and_note_value_only(val: str) -> str:
    s = _norm(val or "")
    s = s.replace("合計量として", "").replace("国際単位", "")
    s = re.sub(r"[gｇ]", "", s)
    return _norm(s)


def has_meaningful_values(rec: dict) -> bool:
    """
    成分名以外の有効値（最大配合量1〜4, 単位, 備考）のどれか1つでも入っているか。
    これが False のときは書き込まない。
    """
    return any(
        (rec.get(k) or "").strip()
        for k in [
            "最大配合量1",
            "最大配合量2",
            "最大配合量3",
            "最大配合量4",
            "単位",
            "備考",
        ]
    )


def row_to_record(row: List[str]) -> Dict[str, str]:
    rec = {
        "成分名": "",
        "最大配合量1": "",
        "最大配合量2": "",
        "最大配合量3": "",
        "最大配合量4": "",
        "単位": "",
        "備考": "",
    }
    if not row:
        return rec
    rec["成分名"] = _norm(row[0])

    if len(row) == 1:
        return rec

    if len(row) == 2:
        clean, unit, note = strip_units_and_note(row[1])
        rec["最大配合量1"] = clean
        rec["単位"] = unit
        rec["備考"] = note
        return rec

    if len(row) == 4:
        clean2, unit, note = strip_units_and_note(row[1])
        rec["最大配合量2"] = clean2
        rec["単位"] = unit
        rec["備考"] = note
        rec["最大配合量3"] = strip_units_and_note_value_only(row[2])
        rec["最大配合量4"] = strip_units_and_note_value_only(row[3])
        return rec

    # その他は 2要素扱いでフォールバック
    clean, unit, note = strip_units_and_note(row[1])
    rec["最大配合量1"] = clean
    rec["単位"] = unit
    rec["備考"] = note
    return rec


# -------------- Google Sheets --------------
def connect_gspread():
    import gspread

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/drive.readonly",
    ]
    return gspread.service_account(filename=SERVICE_ACCOUNT_JSON, scopes=scopes)


def open_or_create_worksheet(gc, spreadsheet_id: str, title: str):
    sh = gc.open_by_key(spreadsheet_id)
    try:
        return sh.worksheet(title)
    except Exception:
        return sh.add_worksheet(title=title, rows=2000, cols=30)


def first_empty_row(ws) -> int:
    # ヘッダーがある前提で、既存最終行の次の行を返す
    values = ws.get_all_values()
    return len(values) + 1


def record_to_row_for_sheet(rec: Dict[str, str], date_str: str, url: str) -> List:
    """
    A..O の並びで返す（未指定列は空文字）
      B: 今日, C:0, D:成分名, F:最大配合量1, H:最大配合量2, I:最大配合量3,
      J:最大配合量4, K:単位, L:備考, O:URL
    """
    return [
        0,  # A 変更フラグ（任意で0）
        date_str,  # B 今日
        0,  # C グループID
        rec.get("成分名", ""),  # D
        "",  # E 規制区分（空）
        rec.get("最大配合量1", ""),  # F
        "",  # G 使用対象・条件（空）
        rec.get("最大配合量2", ""),  # H
        rec.get("最大配合量3", ""),  # I
        rec.get("最大配合量4", ""),  # J
        rec.get("単位", ""),  # K
        rec.get("備考", ""),  # L
        "",  # M 予備
        "",  # N 予備
        url,  # O ソースURL（HTML）
    ]


def append_records_to_sheet(ws, records: List[Dict[str, str]], url: str):
    if not records:
        print("[INFO] 追記対象なし。")
        return
    today_str = datetime.now().strftime("%Y/%m/%d")
    rows = [record_to_row_for_sheet(r, today_str, url) for r in records]
    start = first_empty_row(ws)
    end = start + len(rows) - 1
    ws.update(f"A{start}:O{end}", rows, value_input_option="USER_ENTERED")
    print(f"[DONE] {len(rows)} 行を {ws.title}!A{start}:O{end} に追記しました。")


# -------------- メイン --------------
def main():
    # 1) HTML→3D行列
    tables_rows = build_tables_rows(URL, iframe_first=IFRAME_FIRST)

    # 2) 行→JSON（dict） + フィルタ（成分名しかない行は除外）
    records: List[Dict[str, str]] = []
    for table in tables_rows:
        for row in table:
            rec = row_to_record(row)
            if rec["成分名"] and has_meaningful_values(rec):  # ← ここでフィルタ
                records.append(rec)

    print(f"[INFO] JSONレコード数(書き込み対象): {len(records)}")
    # （確認用のprintは任意）
    for i, r in enumerate(records, 1):
        print(f"{i:03d}: {r}")

    # 3) スプレッドシートへ追記
    gc = connect_gspread()
    ws = open_or_create_worksheet(gc, SPREADSHEET_ID, SHEET_TITLE)
    append_records_to_sheet(ws, records, URL)


if __name__ == "__main__":
    main()

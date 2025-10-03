import io, sys, re, os, json, argparse
from typing import List, Dict, Set
from datetime import datetime
import pandas as pd

# ========= 既定値（引数で上書き可） =========
PDF_URL_DEFAULT = "https://www.mhlw.go.jp/content/000491511.pdf"
PAGES_DEFAULT = "all"             # "all" / "2-12" / "2,4,9" など
EXCLUDE_PAGES: Set[int] = set()   # 例: {1}
USE_AUTO_PAGE_RANGE = False
AUTO_START_PAGE = 2               # True のとき 2-最終 に自動化

# Google スプレッドシート
SERVICE_ACCOUNT_JSON = "./service_account.json"
SPREADSHEET_ID = "1iHKZn9y-AZkeO9TUf0kN3_3GBIthwtTr4i19tNStg8U"
APPEND_SHEET_TITLE = "更新情報一覧"  # 追記先シート名

# 出力オプション（必要なら JSON も保存）
SAVE_JSON_ALL = False             # True にすれば ALL JSON も保存
JSON_DIR = "./json_out"
JSON_ALL_FILENAME = "ccc_ALL.json"

# Camelot（latticeのみ）
LATTICE_LINE_SCALE = 40
LATTICE_COPY_TEXT  = ['h', 'v']
TEXT_STRIP = "\n"

# ========= 引数 =========
def parse_args():
    p = argparse.ArgumentParser(description="PDF→表抽出→整形→スプレッドシート追記")
    p.add_argument("pdf_url", nargs="?", default=PDF_URL_DEFAULT, help="PDFのURL")
    p.add_argument("--pages", default=PAGES_DEFAULT, help='ページ指定（"all", "2-10", "2,4,9" など）')
    return p.parse_args()

# ========= 共通 =========
def download_pdf_bytes(url: str) -> bytes:
    import urllib.request
    with urllib.request.urlopen(url) as r:
        return r.read()

def to_pages_arg(pdf_bytes: bytes, pages_value: str) -> str:
    """--pages 指定 or 既定値を採用。USE_AUTO_PAGE_RANGE=True の場合は 2-最終 などに自動化。"""
    if not USE_AUTO_PAGE_RANGE:
        return pages_value
    from pypdf import PdfReader
    last = len(PdfReader(io.BytesIO(pdf_bytes)).pages)
    start = min(max(1, int(AUTO_START_PAGE)), last)
    return f"{start}-{last}"

def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)

# ========= クリーニング =========
def clean_df(df: pd.DataFrame) -> pd.DataFrame:
    def _strip_cell(x):
        if isinstance(x, str):
            y = x.replace("\u3000", " ")  # 全角→半角スペース

            # 全角括弧の内部スペース削除: "（８ ～ 10 E.O. ）" → "（８～10E.O.）"
            def _rm_spaces_inside_paren(m):
                inner = re.sub(r"[ \t\u3000]+", "", m.group(1))
                return f"（{inner}）"
            y = re.sub(r"（([^）]*)）", _rm_spaces_inside_paren, y)

            # 固定置換
            y = re.sub(r"(合計量として)\s+", r"\1", y)  # 「合計量として」の直後スペース削除
            y = re.sub(r"\s+(?=国際単位)", "", y)        # 「国際単位」の直前スペース削除
            # 「 2－エチル」→「,2－エチル」（直前がカンマでない半角スペースを対象）
            y = re.sub(r"(^|[^,])\s2－エチル", r"\1,2－エチル", y)

            # 余分な空白の正規化
            y = re.sub(r"[ \t]+", " ", y).strip()
            return y
        return x
    return df.map(_strip_cell)

# ========= 行フィルタ =========
DIGIT_ROW_MIN_RATIO = 0.8  # 非空セルのうち「数字だけ」の割合がこの値以上なら削除

def _is_int_string(s: str) -> bool:
    return bool(re.fullmatch(r"\d+", s))

def _norm_no_space(s: str) -> str:
    return re.sub(r"\s+", "", str(s).replace("\u3000", " ")).strip()

def drop_unwanted_rows(df: pd.DataFrame) -> pd.DataFrame:
    """数字だけ行/「成分名」行を削除"""
    if df is None or df.empty:
        return df
    keep_idx = []
    for i, row in df.iterrows():
        cells = [str(v).strip() for v in row.tolist()]
        nonempty = [c for c in cells if c and c.lower() != "nan"]
        digit_only_cnt = sum(1 for c in nonempty if _is_int_string(c))
        is_digit_row = (len(nonempty) > 0) and (digit_only_cnt / len(nonempty) >= DIGIT_ROW_MIN_RATIO)
        first_norm = _norm_no_space(row.iloc[0]) if len(row) > 0 else ""
        is_seibunmei = (first_norm == "成分名")
        if not (is_digit_row or is_seibunmei):
            keep_idx.append(i)
    return df.loc[keep_idx].reset_index(drop=True)

# ========= トークン移動（左列→右列(位置-1)） =========
_TOKEN_IS_AMOUNT = re.compile(r"^\d+(?:\.\d+)?(?:ｇ|国際単位)$")

def _split_tokens(s: str) -> list[str]:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return []
    return [t for t in re.split(r"\s+", str(s).strip()) if t]

def _join_tokens(tokens: list[str]) -> str:
    return " ".join([t for t in tokens if t != ""])

def move_amount_token_from_col1_to_col2(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty or df.shape[1] < 2:
        return df
    out = df.copy()
    for r in range(len(out)):
        c1 = _split_tokens(out.iat[r, 0])
        c2 = _split_tokens(out.iat[r, 1])
        hits = [i for i, tok in enumerate(c1) if _TOKEN_IS_AMOUNT.match(tok)]
        for i_hit in sorted(hits, reverse=True):
            insert_pos = i_hit - 1
            if insert_pos < 0:
                continue
            moved = c1.pop(i_hit)
            if insert_pos > len(c2):
                c2.extend([""] * (insert_pos - len(c2)))
            c2.insert(insert_pos, moved)
        out.iat[r, 0] = _join_tokens(c1)
        out.iat[r, 1] = _join_tokens(c2)
    return out

# ========= テーブル抽出（latticeのみ） =========
def extract_tables_lattice_only(pdf_bytes: bytes, pages: str = "all"):
    import camelot
    items: List[Dict] = []
    kwargs = dict(
        flavor="lattice",
        pages=pages,
        strip_text=TEXT_STRIP,
        line_scale=LATTICE_LINE_SCALE,
        copy_text=LATTICE_COPY_TEXT,
    )
    try:
        tables = camelot.read_pdf(io.BytesIO(pdf_bytes), **kwargs)
        for order, t in enumerate(tables):
            p = int(t.page)
            if p in EXCLUDE_PAGES:
                continue
            cdf = clean_df(t.df)
            if cdf is None or cdf.empty or cdf.shape[0] < 2 or cdf.shape[1] < 2:
                continue
            items.append({"page": p, "order": order, "df": cdf})
    except Exception as e:
        print("[ERROR] lattice 解析で例外:", e, file=sys.stderr)
    items.sort(key=lambda x: (x["page"], x["order"]))
    return items

# ========= JSON化前の追加前処理 =========
def squash_left_when_many_tokens_and_right_one(df: pd.DataFrame) -> pd.DataFrame:
    """
    各行をスペースで分割したとき、
      左列トークン数>=3 かつ 右列トークン数==1 の場合、
      左列を「スペース全削除」して1トークン化する。
    """
    if df is None or df.empty or df.shape[1] < 2:
        return df
    out = df.copy()
    for r in range(len(out)):
        c1 = _split_tokens(out.iat[r, 0])
        c2 = _split_tokens(out.iat[r, 1])
        if len(c1) >= 3 and len(c2) == 1:
            out.iat[r, 0] = "".join(c1)
    return out

# ========= List[dict] 生成（各 dict に url 追加） =========
def _strip_gokei_and_flag(s: str) -> tuple[str, bool]:
    if "合計量として" in (s or ""):
        return s.replace("合計量として", "").strip(), True
    return s or "", False

def _has_kokusai_tanni(*values: str) -> bool:
    return any(("国際単位" in (v or "")) for v in values)

def _strip_units_for_ryou(s: str) -> str:
    s = s or ""
    return re.sub(r"(?:\s*ｇ\s*|\s*国際単位\s*)", "", s).strip()

def _contains_haigou_fuka(s: str) -> bool:
    s = s or ""
    return ("配合負荷" in s) or ("配合不可" in s)

def df_to_records(df: pd.DataFrame, pdf_url: str) -> List[Dict]:
    """
    指定仕様で 1行→1レコードの dict を生成し、各 dict に "url": <pdf_url> を付与。
    - 場合1: 左と右(1列目)のトークン数が等しい
    - 場合2: 左の方が多い → 左0番目を「条件」に切り取ってから処理
    - 2列テーブル: 右1列目を ryou1 に
    - 4列以上: 右の2〜4列を ryou2〜4 に、右側のどこかに「合計量としてX」があれば ryou1=X
    - ryou1〜4 は「ｇ」「国際単位」を除去
    - ryou1 に「配合負荷/配合不可」が含まれる場合は tanni=""（空）
    """
    recs: List[Dict] = []
    if df is None or df.empty:
        return recs

    ncols = df.shape[1]
    use_4cols = ncols >= 4

    for r in range(len(df)):
        c1 = _split_tokens(df.iat[r, 0])
        c2 = _split_tokens(df.iat[r, 1]) if ncols >= 2 else []
        c3 = _split_tokens(df.iat[r, 2]) if ncols >= 3 else []
        c4 = _split_tokens(df.iat[r, 3]) if ncols >= 4 else []

        equal_len = (len(c1) == len(c2))
        cond = ""
        if not equal_len and len(c1) > 0:
            cond = c1.pop(0)

        count = len(c1)
        for i in range(count):
            seibunn = c1[i] if i < len(c1) else ""

            if not use_4cols:
                r1_raw = c2[i] if i < len(c2) else ""
                r1_no_gokei, had_g = _strip_gokei_and_flag(r1_raw)
                r1 = _strip_units_for_ryou(r1_no_gokei)
                tanni = "国際単位" if _has_kokusai_tanni(r1_raw) else "g"
                if _contains_haigou_fuka(r1):
                    tanni = ""
                recs.append({
                    "seibunn": seibunn,
                    "条件": cond if not equal_len else "",
                    "ryou1": r1, "ryou2": "", "ryou3": "", "ryou4": "",
                    "tanni": tanni,
                    "bikou": "合計量として" if had_g else "",
                    "url": pdf_url,
                })
            else:
                v2_raw = c2[i] if i < len(c2) else ""
                v3_raw = c3[i] if i < len(c3) else ""
                v4_raw = c4[i] if i < len(c4) else ""

                r1_raw = ""
                had_g = False
                for cand in (v2_raw, v3_raw, v4_raw):
                    if "合計量として" in cand:
                        r1_raw = cand
                        _, had_g = _strip_gokei_and_flag(cand)
                        break

                r1 = _strip_units_for_ryou(_strip_gokei_and_flag(r1_raw)[0]) if r1_raw else ""
                r2 = _strip_units_for_ryou(v2_raw)
                r3 = _strip_units_for_ryou(v3_raw)
                r4 = _strip_units_for_ryou(v4_raw)

                tanni = "国際単位" if _has_kokusai_tanni(r1_raw, v2_raw, v3_raw, v4_raw) else "g"
                if _contains_haigou_fuka(r1):
                    tanni = ""
                recs.append({
                    "seibunn": seibunn,
                    "条件": cond if not equal_len else "",
                    "ryou1": r1, "ryou2": r2, "ryou3": r3, "ryou4": r4,
                    "tanni": tanni,
                    "bikou": "合計量として" if had_g else "",
                    "url": pdf_url,
                })
    return recs

# ========= Google Sheets =========
def connect_gspread():
    import gspread
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.readonly",
        "https://www.googleapis.com/auth/drive.file",
    ]
    return gspread.service_account(filename=SERVICE_ACCOUNT_JSON, scopes=scopes)

def open_or_create_worksheet(gc, spreadsheet_id: str, title: str):
    sh = gc.open_by_key(spreadsheet_id)
    try:
        return sh.worksheet(title)
    except Exception:
        return sh.add_worksheet(title=title, rows=2000, cols=30)

def first_empty_row(ws) -> int:
    values = ws.get_all_values()
    return len(values) + 1

def record_to_row(rec: Dict, today_str: str) -> List:
    """
    スプレッドシート行データ作成
    A:0, B:日付, C:0, D:seibunn, E:(空), F:ryou1, G:条件, H:ryou2, I:ryou3, J:ryou4, K:tanni, L:bikou,
    M:(空), N:(空), O:url
    """
    seibun = rec.get("seibunn") or rec.get("seibun") or ""
    return [
        0,                           # A: 変更フラグ
        today_str,                   # B: 今日の日付
        0,                           # C: グループID
        seibun,                      # D: 成分名
        "",                          # E: 規制区分（空欄）
        rec.get("ryou1", ""),        # F: 配合上限値（一般）
        rec.get("条件", ""),         # G: 使用対象・条件（一般）
        rec.get("ryou2", ""),        # H:非粘膜・洗い流す上限値
        rec.get("ryou3", ""),        # I:非粘膜・洗い流さない上限値
        rec.get("ryou4", ""),        # J:粘膜用上限値
        rec.get("tanni", ""),        # K:単位
        rec.get("bikou", ""),        # L: 備考
        "",                          # M: 予備
        "",                          # N: 予備
        rec.get("url", ""),          # O: PDF URL
    ]

def append_records_to_sheet(ws, records: List[Dict]):
    if not records:
        print("[INFO] 追記対象なし。")
        return
    today_str = datetime.now().strftime("%Y/%m/%d")
    rows = [record_to_row(r, today_str) for r in records]
    start = first_empty_row(ws)
    end = start + len(rows) - 1
    ws.update(f"A{start}:O{end}", rows, value_input_option="USER_ENTERED")
    print(f"[DONE] {len(rows)} 行を {ws.title}!A{start}:O{end} に追記しました。")

# ========= メイン =========
def main():
    args = parse_args()
    pdf_url = args.pdf_url or PDF_URL_DEFAULT
    pages_opt = args.pages or PAGES_DEFAULT

    print(f"[INFO] Downloading PDF ... {pdf_url}")
    pdf_bytes = download_pdf_bytes(pdf_url)
    pages_arg = to_pages_arg(pdf_bytes, pages_opt)
    print(f"[INFO] Lattice only. pages='{pages_arg}', exclude={sorted(EXCLUDE_PAGES) or '-'}")

    # 1) PDF→表抽出
    items = extract_tables_lattice_only(pdf_bytes, pages_arg)
    if not items:
        print("[ERROR] 表が検出できませんでした。line_scale / --pages / EXCLUDE_PAGES を調整してください。", file=sys.stderr)
        sys.exit(1)

    # 2) 各テーブルを整形しながら List[dict] へ（各 dict に url を付与）
    all_records: List[Dict] = []
    for it in items:
        df = it["df"]
        df = drop_unwanted_rows(df)
        df = move_amount_token_from_col1_to_col2(df)
        df = squash_left_when_many_tokens_and_right_one(df)  # ★ JSON前処理
        recs = df_to_records(df, pdf_url=pdf_url)
        all_records.extend(recs)

    print(f"[INFO] JSON records generated: {len(all_records)}")

    # 3) （任意）ALL JSON を保存
    if SAVE_JSON_ALL:
        ensure_dir(JSON_DIR)
        path_all = os.path.join(JSON_DIR, JSON_ALL_FILENAME)
        with open(path_all, "w", encoding="utf-8") as f:
            json.dump(all_records, f, ensure_ascii=False, indent=2)
        print(f"[INFO] Saved ALL JSON: {path_all}")

    # 4) シートへ追記（O列にURLを書き込む）
    print("[INFO] Appending to Google Sheet ...")
    gc = connect_gspread()
    ws = open_or_create_worksheet(gc, SPREADSHEET_ID, APPEND_SHEET_TITLE)
    append_records_to_sheet(ws, all_records)

    print("[DONE] 完了。")

if __name__ == "__main__":
    main()

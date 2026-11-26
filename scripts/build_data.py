#!/usr/bin/env python3
"""이마트 POS 엑셀(data/raw/*.xlsx)을 대시보드용 data/pos_data.json 으로 변환합니다.

같은 양식의 엑셀이 두 종류 있습니다.
- 수량 파일: 일별 칸에 판매수량
- 금액 파일: 일별 칸에 판매금액(이마트 매입원가 기준)

파일 종류는 ① 파일명 키워드(금액·매출·amount·amt → 금액, 수량·qty → 수량)로 정하고,
키워드가 없으면 ② 값 크기로 자동 판별합니다(0이 아닌 값의 중앙값이 100 이상이면 금액).
파일명에는 연월 6자리(YYYYMM, 예 202501)가 있어야 합니다. 같은 연월·종류 파일이 여러 개면 이름순 마지막 파일을 씁니다.

품목군 규칙은 index.html 의 groupOf() 와 같습니다. 한쪽을 바꾸면 다른 쪽도 같이 바꿔 주세요.

사용법:  python scripts/build_data.py
"""
import collections
import datetime
import glob
import json
import os
import re
import statistics
import sys

from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW_DIR = os.path.join(ROOT, "data", "raw")
OUT_FILE = os.path.join(ROOT, "data", "pos_data.json")

CH_ORDER = ["이마트", "노브랜드", "에브리데이", "트레이더스", "에브리데이B2B"]
GROUPS = ["인스턴트 커피", "원두", "RTD 음료", "액상 에스프레소", "스테비아 커피믹스", "커피믹스",
          "분말 아이스티·레몬", "진액차", "스테비아 차·라떼", "드립백", "필터·용품"]
REQUIRED = ["업태명", "점포코드", "점포명", "상품코드", "상품명"]
DAY_RE = re.compile(r"^(\d{1,2})월(\d{1,2})일$")
YM_RE = re.compile(r"(20\d{2})(0[1-9]|1[0-2])")
RTD_RE = re.compile(r"(\d+(\.\d+)?L|\d{3,}ml)", re.I)  # 1L·2.1L·190ml·400ml 등 (10ml 진액·샷 제외)
AMT_WORDS = ("금액", "매출", "amount", "amt")
QTY_WORDS = ("수량", "qty", "quantity")
KIND_NAME = {"qty": "수량", "amt": "금액"}


def group_of(name: str) -> str:
    n = re.sub(r"\s", "", name or "")
    if "필터" in n:
        return "필터·용품"
    if any(k in n for k in ["진액", "보리차", "수염차", "헛개", "액상스틱"]):
        return "진액차"
    if "에스프레소샷" in n or "리얼에스프레소" in n:
        return "액상 에스프레소"
    if RTD_RE.search(n):
        return "RTD 음료"
    if "아이스티" in n or "레몬즙" in n:
        return "분말 아이스티·레몬"
    if "커피믹스" in n:
        return "스테비아 커피믹스" if "스테비아" in n else "커피믹스"
    if any(k in n for k in ["핫초코", "유자차", "율무차", "쌍화차"]):
        return "스테비아 차·라떼"
    if "드립" in n:
        return "드립백"
    if any(k in n for k in ["까페N블랙", "까페N헤이즐넛", "아이브루", "아이스블랙", "50g", "병커피"]):
        return "인스턴트 커피"
    if re.search(r"\d+입", n) and not any(k in n for k in ["원두", "홀빈", "분쇄", "드립"]):
        return "인스턴트 커피"  # 예: 디카페인헤이즐넛 80입(스틱)
    return "원두"


def text(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return " ".join(str(v).split())


def read_sheet(path):
    """(헤더, 열 위치, 일별 열[(열, 일)], 데이터 행 목록)"""
    wb = load_workbook(path, read_only=True, data_only=True)
    it = wb.worksheets[0].iter_rows(values_only=True)
    hdr = [text(h) for h in next(it)]
    rows = [r for r in it]
    wb.close()
    idx = {h: i for i, h in enumerate(hdr)}
    miss = [h for h in REQUIRED if h not in idx]
    if miss:
        raise ValueError(f"{os.path.basename(path)}: 필요한 열이 없습니다 {miss}")
    days = [(i, DAY_RE.match(h)) for i, h in enumerate(hdr)]
    days = [(i, int(m.group(1)), int(m.group(2))) for i, m in days if m]
    return idx, days, rows


def kind_by_values(days, rows) -> str:
    vals = []
    for r in rows[:3000]:
        for ci, _, _ in days:
            v = r[ci] if ci < len(r) else None
            if isinstance(v, (int, float)) and v > 0:
                vals.append(v)
    return "amt" if vals and statistics.median(vals) >= 100 else "qty"


def kind_of(fname: str, days, rows) -> str:
    # 이마트 원본 파일명은 늘 '기간별매출'로 시작하므로 그 부분은 빼고 판단합니다.
    low = fname.lower().replace("기간별매출", "")
    by_value = kind_by_values(days, rows)
    by_name = "qty" if any(w in low for w in QTY_WORDS) else "amt" if any(w in low for w in AMT_WORDS) else None
    if by_name and by_name != by_value:
        print(f"[주의] {fname}: 파일명은 {KIND_NAME[by_name]}인데 값 크기는 {KIND_NAME[by_value]}처럼 보입니다. 값 크기 기준({KIND_NAME[by_value]})으로 읽습니다.")
    return by_value if by_name != by_value else by_name


def main() -> int:
    files = sorted(f for f in glob.glob(os.path.join(RAW_DIR, "*.xlsx")) if not os.path.basename(f).startswith("~$"))
    if not files:
        print(f"[오류] {RAW_DIR} 에 엑셀 파일이 없습니다.")
        return 1

    picked = {}  # (ym, kind) -> (path, idx, days, rows)
    for f in files:
        name = os.path.basename(f)
        m = YM_RE.search(name)
        if not m:
            print(f"[건너뜀] 파일명에 YYYYMM 이 없습니다: {name}")
            continue
        ym = f"{m.group(1)}-{m.group(2)}"
        idx, days, rows = read_sheet(f)
        for _, mon, _ in days[:1]:
            if mon != int(m.group(2)):
                print(f"[주의] {name}: 일별 열의 월({mon}월)이 파일명({ym})과 다릅니다.")
        kind = kind_of(name, days, rows)
        if (ym, kind) in picked:
            print(f"[주의] {ym} {KIND_NAME[kind]} 파일이 여러 개입니다. {name} 를 씁니다.")
        picked[(ym, kind)] = (f, idx, days, rows)

    months = sorted({ym for ym, _ in picked})
    if not months:
        print("[오류] 읽을 수 있는 파일이 없습니다.")
        return 1
    has = {k: [(ym, k) in picked for ym in months] for k in ("qty", "amt")}

    stores, prods = {}, {}
    sales = collections.defaultdict(lambda: [0.0, 0.0])
    daily = collections.defaultdict(lambda: [0.0, 0.0])
    used = []
    for mi, ym in enumerate(months):
        for k, kind in enumerate(("qty", "amt")):
            if (ym, kind) not in picked:
                continue
            path, idx, days, rows = picked[(ym, kind)]
            total, n_rows = 0.0, 0
            for r in rows:
                st, pc = text(r[idx["점포코드"]]), text(r[idx["상품코드"]])
                if not st or not pc:
                    continue
                ch, name = text(r[idx["업태명"]]), text(r[idx["상품명"]])
                stores[st] = (text(r[idx["점포명"]]), ch)
                prods[pc] = name
                g = group_of(name)
                s = 0.0
                for ci, _, d in days:
                    v = r[ci] if ci < len(r) else None
                    if v:
                        v = float(v)
                        s += v
                        daily[(f"{ym}-{d:02d}", ch, g)][k] += v
                sales[(mi, st, pc)][k] += s
                total += s
                n_rows += 1
            used.append(os.path.basename(path))
            unit = "개" if kind == "qty" else "원"
            print(f"[읽음] {ym} {KIND_NAME[kind]}  {os.path.basename(path)}  행 {n_rows:,}  합계 {total:,.0f}{unit}")

    chans = [c for c in CH_ORDER if any(v[1] == c for v in stores.values())]
    chans += sorted({v[1] for v in stores.values()} - set(chans))
    ci = {c: i for i, c in enumerate(chans)}
    gi = {g: i for i, g in enumerate(GROUPS)}
    st_codes, pr_codes = sorted(stores), sorted(prods)
    si = {c: i for i, c in enumerate(st_codes)}
    pi = {c: i for i, c in enumerate(pr_codes)}

    def num(x):
        return int(x) if float(x).is_integer() else round(x, 3)

    def pair(mi, v):
        return [num(v[0]) if has["qty"][mi] else None, num(v[1]) if has["amt"][mi] else None]

    month_of = {ym: i for i, ym in enumerate(months)}
    out = {
        "version": 2,
        "generated": datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9))).isoformat(timespec="seconds"),
        "source_files": used,
        "months": months,
        "has_qty": has["qty"],
        "has_amt": has["amt"],
        "channels": chans,
        "groups": GROUPS,
        "stores": [[c, stores[c][0], ci[stores[c][1]]] for c in st_codes],
        "products": [[c, prods[c], gi[group_of(prods[c])]] for c in pr_codes],
        # [월 index, 점포 index, 상품 index, 수량, 금액]  (그 달 파일이 없으면 null)
        "sales": [[m, si[s], pi[p]] + pair(m, v) for (m, s, p), v in sorted(sales.items()) if v[0] or v[1]],
        # [일자, 업태 index, 품목군 index, 수량, 금액]
        "daily": [[d, ci[c], gi[g]] + pair(month_of[d[:7]], v) for (d, c, g), v in sorted(daily.items()) if v[0] or v[1]],
    }
    os.makedirs(os.path.dirname(OUT_FILE), exist_ok=True)
    with open(OUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    tq = sum(v[0] for v in sales.values())
    ta = sum(v[1] for v in sales.values())
    am = [ym for ym, h in zip(months, has["amt"]) if h]
    print(f"[완료] {OUT_FILE}  {months[0]}~{months[-1]}  점포 {len(st_codes)}  상품 {len(pr_codes)}")
    print(f"       총 수량 {tq:,.0f}개  총 금액 {ta:,.0f}원  금액 있는 달 {len(am)}개 {am}  ({os.path.getsize(OUT_FILE) / 1e6:.1f}MB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())

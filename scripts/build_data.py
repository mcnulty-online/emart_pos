#!/usr/bin/env python3
"""이마트 POS 엑셀(data/raw/*.xlsx)을 대시보드용 data/pos_data.json 으로 변환합니다.

- 파일명에 YYYYMM(예: 202501)이 들어 있어야 합니다. 같은 월 파일이 여러 개면 이름순으로 마지막 파일을 씁니다.
- 시트 첫 행: 업태명, 점포코드, 점포명, 상품코드, 상품명, …, MM월DD일(일별 수량)…, 합계, 평균
- 품목군 규칙은 index.html 의 groupOf() 와 같습니다. 한쪽을 바꾸면 다른 쪽도 같이 바꿔 주세요.

사용법:  python scripts/build_data.py
"""
import collections
import datetime
import glob
import json
import os
import re
import sys

from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW_DIR = os.path.join(ROOT, "data", "raw")
OUT_FILE = os.path.join(ROOT, "data", "pos_data.json")

CH_ORDER = ["이마트", "노브랜드", "에브리데이", "트레이더스", "에브리데이B2B"]
GROUPS = ["인스턴트 커피", "원두", "RTD 음료", "액상 에스프레소", "스테비아 커피믹스",
          "분말 아이스티·레몬", "진액차", "스테비아 차·라떼", "드립백", "필터·용품"]
REQUIRED = ["업태명", "점포코드", "점포명", "상품코드", "상품명"]
DAY_RE = re.compile(r"^(\d{1,2})월(\d{1,2})일$")
YM_RE = re.compile(r"(20\d{2})(0[1-9]|1[0-2])")
RTD_RE = re.compile(r"(1L|1\.5L|500ml|190ml)", re.I)


def group_of(name: str) -> str:
    n = re.sub(r"\s", "", name or "")
    if "필터" in n:
        return "필터·용품"
    if "진액" in n or "보리차" in n:
        return "진액차"
    if RTD_RE.search(n):
        return "RTD 음료"
    if "아이스티" in n or "레몬즙" in n:
        return "분말 아이스티·레몬"
    if "에스프레소샷" in n or "리얼에스프레소" in n:
        return "액상 에스프레소"
    if "커피믹스" in n:
        return "스테비아 커피믹스"
    if any(k in n for k in ["핫초코", "유자차", "율무차", "쌍화차"]):
        return "스테비아 차·라떼"
    if "드립" in n:
        return "드립백"
    if any(k in n for k in ["까페N블랙", "까페N헤이즐넛", "아이브루", "아이스블랙", "50g", "병커피"]):
        return "인스턴트 커피"
    return "원두"


def text(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return " ".join(str(v).split())


def main() -> int:
    files = sorted(glob.glob(os.path.join(RAW_DIR, "*.xlsx")))
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    if not files:
        print(f"[오류] {RAW_DIR} 에 엑셀 파일이 없습니다.")
        return 1

    by_ym = {}
    for f in files:
        m = YM_RE.search(os.path.basename(f))
        if not m:
            print(f"[건너뜀] 파일명에 YYYYMM 이 없습니다: {os.path.basename(f)}")
            continue
        ym = f"{m.group(1)}-{m.group(2)}"
        if ym in by_ym:
            print(f"[주의] {ym} 파일이 여러 개입니다. {os.path.basename(f)} 를 씁니다.")
        by_ym[ym] = f
    months = sorted(by_ym)
    if not months:
        print("[오류] 읽을 수 있는 파일이 없습니다.")
        return 1

    stores, prods = {}, {}
    sales = collections.defaultdict(float)
    daily = collections.defaultdict(float)
    used = []
    for mi, ym in enumerate(months):
        path = by_ym[ym]
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        it = ws.iter_rows(values_only=True)
        hdr = [text(h) for h in next(it)]
        idx = {h: i for i, h in enumerate(hdr)}
        miss = [h for h in REQUIRED if h not in idx]
        if miss:
            print(f"[오류] {os.path.basename(path)}: 필요한 열이 없습니다 {miss}")
            return 1
        days = []
        for i, h in enumerate(hdr):
            dm = DAY_RE.match(h)
            if dm:
                if int(dm.group(1)) != int(ym[5:]):
                    print(f"[주의] {os.path.basename(path)}: 열 '{h}' 의 월이 파일명({ym})과 다릅니다.")
                days.append((i, int(dm.group(2))))
        n_rows, total = 0, 0.0
        for r in it:
            st = text(r[idx["점포코드"]])
            pc = text(r[idx["상품코드"]])
            if not st or not pc:
                continue
            ch = text(r[idx["업태명"]])
            name = text(r[idx["상품명"]])
            stores[st] = (text(r[idx["점포명"]]), ch)
            prods[pc] = name
            g = group_of(name)
            q_sum = 0.0
            for ci, d in days:
                v = r[ci] if ci < len(r) else None
                if v:
                    v = float(v)
                    q_sum += v
                    daily[(f"{ym}-{d:02d}", ch, g)] += v
            if q_sum:
                sales[(mi, st, pc)] += q_sum
            n_rows += 1
            total += q_sum
        wb.close()
        used.append(os.path.basename(path))
        print(f"[읽음] {ym}  {os.path.basename(path)}  행 {n_rows:,}  수량 {total:,.0f}")

    chans = [c for c in CH_ORDER if any(v[1] == c for v in stores.values())]
    chans += sorted({v[1] for v in stores.values()} - set(chans))
    ci = {c: i for i, c in enumerate(chans)}
    gi = {g: i for i, g in enumerate(GROUPS)}
    st_codes = sorted(stores)
    pr_codes = sorted(prods)
    si = {c: i for i, c in enumerate(st_codes)}
    pi = {c: i for i, c in enumerate(pr_codes)}
    num = lambda x: int(x) if float(x).is_integer() else round(x, 3)

    out = {
        "version": 1,
        "generated": datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9))).isoformat(timespec="seconds"),
        "source_files": used,
        "months": months,
        "channels": chans,
        "groups": GROUPS,
        "stores": [[c, stores[c][0], ci[stores[c][1]]] for c in st_codes],
        "products": [[c, prods[c], gi[group_of(prods[c])]] for c in pr_codes],
        "sales": [[m, si[s], pi[p], num(q)] for (m, s, p), q in sorted(sales.items()) if q],
        "daily": [[d, ci[c], gi[g], num(q)] for (d, c, g), q in sorted(daily.items()) if q],
    }
    os.makedirs(os.path.dirname(OUT_FILE), exist_ok=True)
    with open(OUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    print(f"[완료] {OUT_FILE}  {months[0]}~{months[-1]}  점포 {len(st_codes)}  상품 {len(pr_codes)}  "
          f"총 수량 {sum(sales.values()):,.0f}  ({os.path.getsize(OUT_FILE) / 1e6:.1f}MB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())

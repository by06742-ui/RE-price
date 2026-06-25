# -*- coding: utf-8 -*-
"""
RE_price (웹앱) — 서울 빌라(연립다세대) 신축 부지가치 산출기
경량 회귀-크리깅: 노후도(사용연수) 회귀로 '신축(0년) 환산' → 법정동 경계 안 공간보간.
빌라+단독 통합(직거래 제외). 가산 가중치 없음(이중계산 방지).
같은 폴더 파일: prices.json, PNU_coords.npz, dongnames.csv
필요 라이브러리: streamlit, numpy, openpyxl
"""
import os, csv, glob, json, io
import numpy as np
import streamlit as st
from openpyxl import Workbook

PYEONG = 3.305785
IDW_K = 8
IDW_P = 2.0
MIN_DONG_PTS = 3
HERE = os.path.dirname(os.path.abspath(__file__))


def _load(here):
    z = np.load(glob.glob(os.path.join(here, "*.npz"))[0], allow_pickle=False)
    pnu = z["pnu"].astype("U19")
    x = z["x"].astype(np.float64); y = z["y"].astype(np.float64)
    area = z["area"].astype(np.float64) if "area" in z.files else np.full(len(pnu), np.nan)
    o = np.argsort(pnu); pnu, x, y, area = pnu[o], x[o], y[o], area[o]
    prices = json.load(open(os.path.join(here, "prices.json"), encoding="utf-8"))
    dn = {}
    with open(os.path.join(here, "dongnames.csv"), encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            c = (row.get("동코드") or "").strip(); n = (row.get("동명") or "").strip()
            if c and n: dn[c] = n
    gun = {}
    for c, n in dn.items():
        g = c[:5]
        if g not in gun: gun[g] = n.split()[0] if n else ""
    kxx, kxy, kv, kdong = [], [], [], []
    for p, price in prices["parcel"].items():
        i = np.searchsorted(pnu, p)
        if i < len(pnu) and pnu[i] == p:
            kxx.append(x[i]); kxy.append(y[i]); kv.append(price); kdong.append(p[:10])
    kxx = np.asarray(kxx, float); kxy = np.asarray(kxy, float); kv = np.asarray(kv, float)
    dong_idx = {}
    for idx, d in enumerate(kdong):
        dong_idx.setdefault(d, []).append(idx)
    for d in dong_idx:
        dong_idx[d] = np.asarray(dong_idx[d], int)
    return {"pnu": pnu, "x": x, "y": y, "area": area, "prices": prices,
            "dn": dn, "gun": gun, "rev": {n: c for c, n in dn.items()},
            "kxx": kxx, "kxy": kxy, "kv": kv, "dong_idx": dong_idx,
            "age_model": prices.get("age_model", {"b1": 0.0, "b2": 0.0})}


@st.cache_resource(show_spinner="데이터 불러오는 중 ...")
def load_data():
    return _load(HERE)


def normalize_pnu(s):
    return "".join(ch for ch in str(s).strip().replace("-", "").replace(" ", "") if ch.isdigit())


def jibun_from_pnu(p, dn):
    bon, bu = int(p[11:15]), int(p[15:19])
    번지 = f"{bon}" if bu == 0 else f"{bon}-{bu}"
    동 = dn.get(p[:10], "")
    return (동 + " " + 번지).strip() if 동 else 번지


def address_to_pnu(addr, rev):
    if not addr: return None
    toks = [t for t in addr.replace(",", " ").replace("\t", " ").split()
            if t not in ("서울특별시", "서울시", "서울")]
    if len(toks) < 2: return None
    번지 = toks[-1]; 동 = " ".join(toks[:-1])
    code = rev.get(동)
    if not code: return None
    san = "2" if 번지.startswith("산") else "1"
    번지 = 번지.lstrip("산")
    bon, bu = (번지.split("-") + ["0"])[:2] if "-" in 번지 else (번지, "0")
    try:
        bon, bu = int(bon), int(bu)
    except ValueError:
        return None
    return f"{code}{san}{bon:04d}{bu:04d}"


def _idw_in_dong(D, qx, qy, dong_code):
    cand = D["dong_idx"].get(dong_code)
    if cand is None or len(cand) < MIN_DONG_PTS:
        return None
    dx = D["kxx"][cand] - qx; dy = D["kxy"][cand] - qy
    d2 = dx * dx + dy * dy
    k = min(IDW_K, len(d2))
    sel = np.argpartition(d2, k - 1)[:k] if k < len(d2) else np.arange(len(d2))
    d = np.sqrt(d2[sel]); o = np.argsort(d); d = d[o]; sel = sel[o]
    if d[0] == 0:
        return float(D["kv"][cand][sel[0]]), 0.0
    w = 1.0 / (d ** IDW_P)
    return float((w * D["kv"][cand][sel]).sum() / w.sum()), float(d[0])


def estimate_one(pnu, D, age=None):
    P = D["prices"]
    p = normalize_pnu(pnu)
    if len(p) != 19:
        return {"PNU": pnu, "지번": "", "토지면적_㎡": None, "토지면적_평": None,
                "신축_평당가_만원": None, "사용연수": age, "대지㎡당가_만원": None,
                "대지평당가_만원": None, "추정방식": "오류(19자리 아님)", "참고": ""}
    b1, b2 = D["age_model"]["b1"], D["age_model"]["b2"]
    factor = lambda a: float(np.exp(b1 * a + b2 * a * a))
    지번 = jibun_from_pnu(p, D["dn"])
    i = np.searchsorted(D["pnu"], p)
    has = i < len(D["pnu"]) and D["pnu"][i] == p
    면적 = round(float(D["area"][i]), 1) if has and not np.isnan(D["area"][i]) else None
    면적평 = round(면적 / PYEONG, 1) if 면적 is not None else None

    def out(base_val, 방식, 참고):
        신축 = round(base_val, 1)
        if age:
            보임 = round(신축 * factor(age), 1); 참고2 = 참고 + f" · 사용 {age}년 환산"
        else:
            보임 = 신축; 참고2 = 참고
        return {"PNU": p, "지번": 지번, "토지면적_㎡": 면적, "토지면적_평": 면적평,
                "신축_평당가_만원": 신축, "사용연수": age if age else 0,
                "대지㎡당가_만원": round(보임 / PYEONG, 1), "대지평당가_만원": 보임,
                "추정방식": 방식, "참고": 참고2}

    if p in P["parcel"]:
        return out(P["parcel"][p], "실측", "실거래")
    if has:
        r = _idw_in_dong(D, D["x"][i], D["y"][i], p[:10])
        if r is not None:
            v, nd = r
            return out(v, "공간보간(동내)", f"최근접 {nd:.0f}m")
    if p[:10] in P["dong"]:
        pr, cnt = P["dong"][p[:10]]
        return out(pr, "동 추정", f"동 {cnt}건")
    if p[:5] in P["gu"]:
        pr, cnt = P["gu"][p[:5]]
        return out(pr, "구 추정", f"구 {cnt}건")
    return out(P["total_med"], "전체 추정", "")


COLS = ["PNU", "지번", "토지면적_㎡", "토지면적_평", "신축_평당가_만원",
        "대지평당가_만원", "대지㎡당가_만원", "사용연수", "추정방식", "참고"]


def rows_to_xlsx(rows):
    wb = Workbook(); ws = wb.active; ws.title = "RE_price"
    ws.append(COLS)
    for r in rows:
        ws.append([r.get(c) for c in COLS])
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()


st.set_page_config(page_title="RE_price 빌라 신축 부지가치", page_icon="🏠", layout="centered")
st.title("🏠 서울 빌라(연립다세대) 신축 부지가치 산출기")
st.caption("대지권(토지지분) 기준 = 물건금액 ÷ 대지권면적. 노후도(사용연수) 회귀로 '신축(0년) 환산' 후 "
           "법정동 경계 안 실거래로 공간보간(하천·대로 미교차). 빌라+단독 통합·직거래 제외. 가산 가중치 없음.")

D = load_data()
st.success(f"준비 완료 — 실측(빌라) {len(D['prices']['parcel']):,}지번 · 좌표 {len(D['pnu']):,}필지")

age = st.number_input("사용연수 (0 = 신축 기준 부지가치, 입력 시 해당 연식으로 감가)",
                      min_value=0, max_value=60, value=0, step=1)

tab1, tab2, tab3 = st.tabs(["① PNU 조회", "② 지번(주소) 조회", "③ 여러 개(엑셀 붙여넣기)"])


def show(r):
    if r["대지평당가_만원"] is None:
        st.warning(r["추정방식"]); return
    면적 = (f"{r['토지면적_㎡']:,.1f}㎡ ({r['토지면적_평']:,.1f}평)"
          if r["토지면적_㎡"] is not None else "면적정보없음")
    c1, c2 = st.columns(2)
    c1.metric((r["지번"] or r["PNU"]) + " · 신축환산 평당가", f"{r['신축_평당가_만원']:,.0f} 만원/평")
    c2.metric("신축환산 ㎡당가", f"{r['신축_평당가_만원']/PYEONG:,.0f} 만원/㎡")
    if r["사용연수"]:
        d1, d2 = st.columns(2)
        d1.metric(f"사용 {r['사용연수']}년 평당가", f"{r['대지평당가_만원']:,.0f} 만원/평")
        d2.metric(f"사용 {r['사용연수']}년 ㎡당가", f"{r['대지㎡당가_만원']:,.0f} 만원/㎡")
    st.write(f"필지면적 {면적} · 방식 {r['추정방식']} ({r['참고']})")


with tab1:
    pnu = st.text_input("PNU(19자리)", placeholder="예: 1111010200100380007")
    if st.button("조회", key="b1") and pnu.strip():
        show(estimate_one(pnu, D, age))

with tab2:
    addr = st.text_input("지번 주소", placeholder="예: 강남구 역삼동 601-1")
    if st.button("조회", key="b2") and addr.strip():
        p = address_to_pnu(addr, D["rev"])
        if not p:
            st.warning("주소를 PNU로 못 바꿨습니다. '자치구 동 번지' 형식인지 확인하세요.")
        else:
            show(estimate_one(p, D, age))

with tab3:
    st.write("엑셀에서 PNU가 세로로 나열된 칸을 복사해 붙여넣으세요. (위 사용연수 입력값이 함께 적용)")
    text = st.text_area("PNU 목록", height=160, placeholder="1111010200100380007\n1168010100106010001\n...")
    if st.button("조회", key="b3"):
        pnus = []
        for line in text.splitlines():
            for tok in line.replace(",", "\t").split("\t"):
                t = normalize_pnu(tok)
                if t: pnus.append(t)
        if not pnus:
            st.info("PNU를 붙여넣으세요.")
        else:
            rows = [estimate_one(p, D, age) for p in pnus]
            st.dataframe(rows, use_container_width=True)
            st.download_button("⬇️ 엑셀(.xlsx)로 저장", data=rows_to_xlsx(rows),
                               file_name="RE_price_result.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.divider()
st.caption("추정 순서: 실측(빌라) → 공간보간(동내) → 동 추정 → 구 추정 → 전체. "
           "단독다가구는 지번 비식별로 동/구 통계에만 반영.")

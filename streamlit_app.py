# -*- coding: utf-8 -*-
"""
RE_price (웹앱) — 서울 빌라(연립다세대) 신축 평당가 산출기
------------------------------------------------------------------
Streamlit 웹앱입니다. Streamlit Community Cloud(무료)에 올리면
브라우저에서 URL로 접속해 사용할 수 있습니다.

같은 폴더에 아래 데이터가 함께 있어야 합니다:
  prices.json     : PNU/동/구별 평당가(미리 계산)
  PNU_coords.npz  : 필지 좌표 + 토지면적
  dongnames.csv   : 법정동코드→동이름
필요 라이브러리: streamlit, numpy, openpyxl
"""

import os
import csv
import glob
import json
import io

import numpy as np
import streamlit as st
from openpyxl import Workbook

PYEONG = 3.305785
IDW_K = 8
IDW_P = 2.0
HERE = os.path.dirname(os.path.abspath(__file__))


# ====================== 데이터 로드 (캐시) ======================
def _load_data(here):
    npz_path = glob.glob(os.path.join(here, "*.npz"))[0]
    z = np.load(npz_path, allow_pickle=False)
    pnu = z["pnu"].astype("U19")
    xy = np.column_stack([z["x"].astype(np.float64), z["y"].astype(np.float64)])
    area = z["area"].astype(np.float64) if "area" in z.files else np.full(len(pnu), np.nan)
    order = np.argsort(pnu)
    pnu, xy, area = pnu[order], xy[order], area[order]

    prices = json.load(open(os.path.join(here, "prices.json"), encoding="utf-8"))

    dn = {}
    with open(os.path.join(here, "dongnames.csv"), encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            c = (row.get("동코드") or "").strip()
            n = (row.get("동명") or "").strip()
            if c and n:
                dn[c] = n
    gun = {}
    for c, n in dn.items():
        g = c[:5]
        if g not in gun:
            gun[g] = n.split()[0] if n else ""

    # 공간보간 학습점(실거래 좌표) 구성
    kx, kv = [], []
    for p, price in prices["parcel"].items():
        i = np.searchsorted(pnu, p)
        if i < len(pnu) and pnu[i] == p:
            kx.append(xy[i]); kv.append(price)
    return {
        "pnu": pnu, "xy": xy, "area": area, "prices": prices,
        "dn": dn, "gun": gun, "rev": {n: c for c, n in dn.items()},
        "kx": np.asarray(kx, float), "kv": np.asarray(kv, float),
    }


@st.cache_resource(show_spinner="데이터 불러오는 중 ...")
def load_data():
    return _load_data(HERE)


# ====================== 추정 로직 ======================
def normalize_pnu(s):
    s = str(s).strip().replace("-", "").replace(" ", "")
    return "".join(ch for ch in s if ch.isdigit())


def jibun_from_pnu(p, dn):
    bon, bu = int(p[11:15]), int(p[15:19])
    번지 = f"{bon}" if bu == 0 else f"{bon}-{bu}"
    동 = dn.get(p[:10], "")
    return (동 + " " + 번지).strip() if 동 else 번지


def address_to_pnu(addr, rev):
    if not addr:
        return None
    toks = [t for t in addr.replace(",", " ").split() if t not in ("서울특별시", "서울시", "서울")]
    if len(toks) < 2:
        return None
    번지, 동 = toks[-1], " ".join(toks[:-1])
    code = rev.get(동)
    if not code:
        return None
    san = "2" if 번지.startswith("산") else "1"
    번지 = 번지.lstrip("산")
    bon, bu = (번지.split("-") + ["0"])[:2] if "-" in 번지 else (번지, "0")
    try:
        bon, bu = int(bon), int(bu)
    except ValueError:
        return None
    return f"{code}{san}{bon:04d}{bu:04d}"


def _idw(kx, kv, xy):
    diff = kx - xy
    d2 = np.einsum("ij,ij->i", diff, diff)
    k = min(IDW_K, len(d2))
    idx = np.argpartition(d2, k - 1)[:k] if k < len(d2) else np.arange(len(d2))
    d = np.sqrt(d2[idx]); o = np.argsort(d); d = d[o]; idx = idx[o]
    if d[0] == 0:
        return float(kv[idx[0]]), 0.0
    w = 1.0 / (d ** IDW_P)
    return float((w * kv[idx]).sum() / w.sum()), float(d[0])


def estimate_one(pnu, D):
    p = normalize_pnu(pnu)
    if len(p) != 19:
        return {"PNU": pnu, "지번": "", "토지면적_㎡": None, "토지면적_평": None, "평당가_만원": None,
                "추정방식": "오류(19자리 아님)", "참고": ""}
    지번 = jibun_from_pnu(p, D["dn"])
    i = np.searchsorted(D["pnu"], p)
    has = i < len(D["pnu"]) and D["pnu"][i] == p
    면적 = round(float(D["area"][i]), 1) if has and not np.isnan(D["area"][i]) else None
    면적평 = round(면적 / PYEONG, 1) if 면적 is not None else None

    def out(v, 방식, 참고):
        return {"PNU": p, "지번": 지번, "토지면적_㎡": 면적, "토지면적_평": 면적평, "평당가_만원": v,
                "추정방식": 방식, "참고": 참고}

    if p in D["prices"]["parcel"]:
        return out(round(D["prices"]["parcel"][p], 1), "실측", "실거래")
    if has and len(D["kx"]):
        v, nd = _idw(D["kx"], D["kv"], D["xy"][i])
        return out(round(v, 1), "공간보간(IDW)", f"최근접거리 {nd:.0f}m")
    if p[:10] in D["prices"]["dong"]:
        pr, cnt = D["prices"]["dong"][p[:10]]
        return out(round(pr, 1), "동 추정", f"동 {cnt}지번")
    if p[:5] in D["prices"]["gu"]:
        pr, cnt = D["prices"]["gu"][p[:5]]
        return out(round(pr, 1), "구 추정", f"구 {cnt}지번")
    return out(round(D["prices"]["total_med"], 1), "전체 추정", "")


def rows_to_xlsx(rows):
    cols = ["PNU", "지번", "토지면적_㎡", "토지면적_평", "평당가_만원", "추정방식", "참고"]
    wb = Workbook(); ws = wb.active; ws.title = "RE_price"
    ws.append(cols)
    for r in rows:
        ws.append([r.get(c) for c in cols])
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()


# ====================== 웹 화면 ======================
st.set_page_config(page_title="RE_price 빌라 평당가", page_icon="🏠", layout="centered")
st.title("🏠 서울 빌라(연립다세대) 신축 평당가 산출기")
st.caption("서울시 실거래가 중 '연립다세대(빌라)'만으로, 특정 지번에 빌라 신축 시 "
           "예상 평당가를 실거래·공간보간(IDW)으로 추정합니다.")
st.warning("⚠️ **초기 버전**입니다. **2026년 현재의 평단가와 상이할 수 있습니다.** 참고용으로만 활용하세요.")

D = load_data()
st.success(f"준비 완료 — 실측 {len(D['prices']['parcel']):,}지번 · 좌표 {len(D['pnu']):,}필지")

tab1, tab2, tab3 = st.tabs(["① PNU 조회", "② 지번(주소) 조회", "③ 여러 개(엑셀 붙여넣기)"])

with tab1:
    pnu = st.text_input("PNU(19자리)", placeholder="예: 1111010200100380007")
    if st.button("조회", key="b1") and pnu.strip():
        r = estimate_one(pnu, D)
        if r["평당가_만원"] is None:
            st.warning(r["추정방식"])
        else:
            면적 = (f"{r['토지면적_㎡']:,.1f}㎡ ({r['토지면적_평']:,.1f}평)"
                  if r["토지면적_㎡"] is not None else "면적정보없음")
            st.metric(r["지번"] or r["PNU"], f"{r['평당가_만원']:,.0f} 만원/평")
            st.write(f"토지면적 {면적} · 방식 {r['추정방식']} ({r['참고']})")

with tab2:
    addr = st.text_input("지번 주소", placeholder="예: 강남구 역삼동 601-1")
    if st.button("조회", key="b2") and addr.strip():
        p = address_to_pnu(addr, D["rev"])
        if not p:
            st.warning("주소를 PNU로 못 바꿨습니다. '자치구 동 번지' 형식인지 확인하세요.")
        else:
            r = estimate_one(p, D)
            면적 = (f"{r['토지면적_㎡']:,.1f}㎡ ({r['토지면적_평']:,.1f}평)"
                  if r["토지면적_㎡"] is not None else "면적정보없음")
            st.metric(r["지번"] or r["PNU"], f"{r['평당가_만원']:,.0f} 만원/평")
            st.write(f"토지면적 {면적} · 방식 {r['추정방식']} ({r['참고']})")

with tab3:
    st.write("엑셀에서 PNU가 세로로 나열된 칸을 복사해 아래에 붙여넣으세요.")
    text = st.text_area("PNU 목록", height=160, placeholder="1111010200100380007\n1168010100106010001\n...")
    if st.button("조회", key="b3"):
        pnus = []
        for line in text.splitlines():
            for tok in line.replace(",", "\t").split("\t"):
                t = normalize_pnu(tok)
                if t:
                    pnus.append(t)
        if not pnus:
            st.info("PNU를 붙여넣으세요.")
        else:
            rows = [estimate_one(p, D) for p in pnus]
            st.dataframe(rows, use_container_width=True)
            st.download_button("⬇️ 엑셀(.xlsx)로 저장", data=rows_to_xlsx(rows),
                               file_name="RE_price_result.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.divider()
st.caption("추정 순서: ① 실측 → ② 공간보간(IDW) → ③ 동/구 추정 → ④ 전체. "
           "토지면적은 연속지적도 필지(대지) 면적입니다.")

# -*- coding: utf-8 -*-
"""
======================================================================
 RE_price (배포용 경량판) — 서울 빌라(연립다세대) 토지가치 산출기
======================================================================
 · 배포용으로 용량을 크게 줄인 버전입니다.
   (무거운 pandas·scipy를 빼고 numpy·openpyxl만 사용 → EXE가 훨씬 가벼움)
 · 실거래가는 미리 계산한 'prices.json'으로, 좌표·면적은 'PNU_coords.npz'로
   내장됩니다. 원본 CSV는 들어가지 않아 더 가볍습니다.
 · 연립다세대(빌라) 실거래만으로 산출합니다.

 [내장 데이터]
   prices.json     : PNU/동/구별 평당가(미리 계산됨)
   PNU_coords.npz  : 필지 중심좌표 + 토지면적 (연속지적도에서 추출)
   dongnames.csv   : 법정동코드→동이름 (지번 표시용)

 [추정 순서]  ① 실측 → ② 공간보간(IDW) → ③ 동/구 추정 → ④ 전체
 필요 라이브러리: numpy, openpyxl  (tkinter는 파이썬 기본 포함)
======================================================================
"""

import os
import sys
import csv
import glob
import json
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import numpy as np
from openpyxl import Workbook

PYEONG = 3.305785
IDW_K = 8
IDW_P = 2.0
MIN_DONG_PTS = 3        # 법정동 내 실거래가 이 개수 이상일 때만 동내 공간보간(경계 미교차)


# ====================================================================
#  내장 파일 찾기 (여러 위치 + 확장자 기반 → 위치/이름에 안 탐)
# ====================================================================
def base_dirs():
    dirs = []
    if hasattr(sys, "_MEIPASS"):
        dirs.append(sys._MEIPASS)
    if getattr(sys, "frozen", False):
        dirs.append(os.path.dirname(sys.executable))
    else:
        dirs.append(os.path.dirname(os.path.abspath(__file__)))
    dirs.append(os.getcwd())
    out, seen = [], set()
    for d in dirs:
        if d and d not in seen:
            seen.add(d); out.append(d)
    return out


def _search_dirs(extra=None):
    dirs = []
    if extra:
        dirs += [extra, os.path.join(extra, "data")]
    for b in base_dirs():
        dirs += [b, os.path.join(b, "data")]
    out, seen = [], set()
    for d in dirs:
        if d and d not in seen:
            seen.add(d); out.append(d)
    return out


def _find_by_ext(ext, extra=None):
    for d in _search_dirs(extra):
        hits = sorted(glob.glob(os.path.join(d, ext)))
        if hits:
            return hits[0]
    return None


def _find_named(name, extra=None):
    for d in _search_dirs(extra):
        p = os.path.join(d, name)
        if os.path.isfile(p):
            return p
    return None


# ====================================================================
#  데이터 로드
# ====================================================================
def _find_coords_npz(extra=None):
    import glob as _g
    cands = []
    dirs = ([extra] if extra else []) + [os.path.dirname(os.path.abspath(sys.argv[0])), os.getcwd()]
    try:
        dirs.append(os.path.dirname(os.path.abspath(__file__)))
    except Exception:
        pass
    for d in dirs:
        if d and os.path.isdir(d):
            cands += _g.glob(os.path.join(d, "PNU_coords*.npz"))
    return cands[0] if cands else None


def load_coords(log, extra=None):
    path = _find_coords_npz(extra) or _find_by_ext("PNU_coords*.npz", extra)
    if not path:
        raise FileNotFoundError("좌표표(PNU_coords*.npz)를 찾지 못했습니다.")
    log(f"필지 좌표·면적 불러오는 중 ... ({os.path.basename(path)})")
    z = np.load(path, allow_pickle=False)
    pnu = z["pnu"].astype("U19")
    xy = np.column_stack([z["x"].astype(np.float64), z["y"].astype(np.float64)])
    area = z["area"].astype(np.float64) if "area" in z.files else np.full(len(pnu), np.nan)
    age = z["age"].astype(np.int16) if "age" in z.files else np.full(len(pnu), -1, np.int16)
    struct = z["struct"].astype(np.uint8) if "struct" in z.files else np.zeros(len(pnu), np.uint8)
    uq = z["uq"].astype(np.uint8) if "uq" in z.files else np.zeros(len(pnu), np.uint8)
    order = np.argsort(pnu)
    log(f"  좌표 보유 필지 {len(pnu):,}개")
    return pnu[order], xy[order], area[order], age[order], struct[order], uq[order]


def load_prices(log, extra=None):
    path = _find_named("prices.json", extra)
    if not path:
        raise FileNotFoundError("가격표(prices.json)를 찾지 못했습니다.")
    log(f"가격표 불러오는 중 ... ({os.path.basename(path)})")
    d = json.load(open(path, encoding="utf-8"))
    log(f"  실측 지번 {len(d['parcel']):,} / 동 {len(d['dong'])} / 구 {len(d['gu'])}")
    return d


def load_ho(log, extra=None):
    # 경량(ho_lite.npz) 우선
    p = _find_named("ho_lite.npz", extra)
    if p:
        log(f"호별 대지권(경량) 불러오는 중 ... ({os.path.basename(p)})")
        z = np.load(p, allow_pickle=False)
        d = {}
        for pnu, hm, fl, ar in zip(z["pnu"].astype("U19"), z["ho"], z["fl"], z["area"]):
            d.setdefault(str(pnu), []).append([str(hm), str(fl), float(ar)])
        log(f"  호데이터 {len(d):,}필지")
        return d
    path = _find_named("ho.json", extra)
    if not path:
        log("  (호 데이터 없음 — 호별 산출 생략)"); return {}
    log(f"호별 대지권 불러오는 중 ... ({os.path.basename(path)})")
    d = json.load(open(path, encoding="utf-8"))
    log(f"  호데이터 {len(d):,}필지")
    return d


class RoadMap:
    """도로명→PNU 조회. 경량(npz, 이진탐색) 또는 dict(json) 모두 지원."""
    def __init__(self, keys=None, pnu=None, d=None):
        self.keys = keys; self.pnu = pnu; self.d = d

    def get(self, key):
        if self.d is not None:
            return self.d.get(key)
        if self.keys is None or len(self.keys) == 0:
            return None
        i = int(np.searchsorted(self.keys, key))
        if 0 <= i < len(self.keys) and self.keys[i] == key:
            return str(self.pnu[i])
        return None

    def __contains__(self, key):
        return self.get(key) is not None

    def __getitem__(self, key):
        return self.get(key)

    def __len__(self):
        if self.d is not None:
            return len(self.d)
        return len(self.keys) if self.keys is not None else 0


def load_road(log, extra=None):
    p = _find_named("road_lite.npz", extra)
    if p:
        log(f"도로명 주소(경량) 불러오는 중 ... ({os.path.basename(p)})")
        z = np.load(p, allow_pickle=False)
        rm = RoadMap(keys=z["keys"], pnu=z["pnu"].astype("U19"))
        log(f"  도로명 {len(rm):,}개"); return rm
    path = _find_named("road.json", extra)
    if not path:
        log("  (도로명 데이터 없음 — 도로명 검색 생략)"); return RoadMap(d={})
    log(f"도로명 주소 불러오는 중 ... ({os.path.basename(path)})")
    rm = RoadMap(d=json.load(open(path, encoding="utf-8")))
    log(f"  도로명 {len(rm):,}개"); return rm


# 층별 보정 (동 대비): 반지하 0.625 · 1층 0.854 · 2~3층 1.0(기준) · 4층↑ 1.093
FLOOR_MULT = {"B": 0.625, "1": 0.854, "2": 1.0, "4": 1.093}
FLOOR_LABEL = {"B": "반지하", "1": "1층", "2": "2~3층", "4": "4층↑"}
FLOOR_ORDER = {"반지하": 0, "1층": 1, "2~3층": 2, "4층↑": 3}


def load_dongnames(log, extra=None):
    path = _find_named("dongnames.csv", extra)
    if not path:
        log("  (dongnames.csv 없음 — 지번은 번지만 표시)")
        return {}, {}
    dn = {}
    with open(path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            code = (row.get("동코드") or "").strip()
            name = (row.get("동명") or "").strip()
            if code and name:
                dn[code] = name
    gu = {}
    for code, name in dn.items():
        g = code[:5]
        if g not in gu:
            gu[g] = name.split()[0] if name else ""
    log(f"  동 이름표 {len(dn):,}개")
    return dn, gu


# ====================================================================
#  좌표 검색 + 공간보간(IDW) — scipy 없이 numpy로 직접 구현
# ====================================================================
def lookup_idx(pnu_sorted, p):
    i = np.searchsorted(pnu_sorted, p)
    if i < len(pnu_sorted) and pnu_sorted[i] == p:
        return int(i)
    return -1


class BarrierKNN:
    """법정동 내부의 실거래만으로 IDW 추정.
    법정동 경계는 대개 하천·대로를 따라가므로, 동 내부로 제한하면
    공간보간이 하천/대로 경계를 넘어 영향을 주지 않습니다."""
    def __init__(self, pts, vals, dong_codes):
        self.pts = np.asarray(pts, float)
        self.vals = np.asarray(vals, float)
        self.dong_idx = {}
        for i, d in enumerate(dong_codes):
            self.dong_idx.setdefault(d, []).append(i)
        for d in self.dong_idx:
            self.dong_idx[d] = np.asarray(self.dong_idx[d], int)

    def idw_in_dong(self, xy, dong_code, k=IDW_K, p=IDW_P, min_pts=MIN_DONG_PTS):
        cand = self.dong_idx.get(dong_code)
        if cand is None or len(cand) < min_pts:
            return None   # 동 내 표본 부족 → 상위 단계(동/구 중앙값)로
        pts = self.pts[cand]; vals = self.vals[cand]
        diff = pts - xy
        d2 = np.einsum("ij,ij->i", diff, diff)
        kk = min(k, len(d2))
        idx = np.argpartition(d2, kk - 1)[:kk] if kk < len(d2) else np.arange(len(d2))
        d = np.sqrt(d2[idx])
        o = np.argsort(d); d = d[o]; idx = idx[o]
        if d[0] == 0:
            return vals[idx[0]].tolist(), 0.0
        w = 1.0 / (d ** p)
        wv = (w[:, None] * vals[idx]).sum(0) / w.sum()
        return wv.tolist(), float(d[0])


def build_spatial(parcel_prices, pnu_sorted, xy_sorted, log):
    log("공간보간 학습점 구성 중 (법정동 경계 기준) ...")
    kx, kv, kd = [], [], []
    for pnu, price in parcel_prices.items():
        i = lookup_idx(pnu_sorted, pnu)
        if i >= 0:
            kx.append(xy_sorted[i]); kv.append(price); kd.append(pnu[:10])
    log(f"  학습점 {len(kx):,}개")
    return BarrierKNN(kx, kv, kd)


# ====================================================================
#  PNU/지번 ↔ 추정
# ====================================================================
def normalize_pnu(s):
    s = str(s).strip().replace("-", "").replace(" ", "")
    return "".join(ch for ch in s if ch.isdigit())


def jibun_from_pnu(p, dong_names):
    bon, bu = int(p[11:15]), int(p[15:19])
    번지 = f"{bon}" if bu == 0 else f"{bon}-{bu}"
    동명 = dong_names.get(p[:10], "")
    return (동명 + " " + 번지).strip() if 동명 else 번지


def address_to_pnu(addr, dong_rev):
    if not addr:
        return None
    toks = addr.replace(",", " ").replace("\t", " ").split()
    toks = [t for t in toks if t not in ("서울특별시", "서울시", "서울")]
    if len(toks) < 2:
        return None
    번지 = toks[-1]
    동명 = " ".join(toks[:-1])
    code = dong_rev.get(동명)
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


def norm_road(s):
    """도로명주소 정규화: '서울/서울특별시' 및 괄호(법정동) 제거, 공백 정리"""
    import re
    s = re.sub(r"\(.*?\)", "", str(s).strip())
    for k in ("서울특별시", "서울시", "서울"):
        if s.startswith(k):
            s = s[len(k):]
    return re.sub(r"\s+", " ", s).strip()


def resolve_address(addr, dong_rev, road):
    """지번이면 지번 변환, 도로명이면 road 맵에서 PNU 조회 (둘 다 시도)."""
    if not addr:
        return None
    # 도로명(로/길/대로 포함) 우선 시도
    key = norm_road(addr)
    if road and key in road:
        return road[key]
    p = address_to_pnu(addr, dong_rev)
    if p:
        return p
    return road.get(key) if road else None


def estimate_one(pnu, ctx, age=None):
    (parcel, dong, gu, total, pnu_sorted, xy_sorted, area_sorted,
     dong_names, gu_names, knn, building, age_sorted, struct_sorted, uq_sorted) = ctx
    p = normalize_pnu(pnu)
    if len(p) != 19:
        return {"PNU": pnu, "지번": "", "토지면적_㎡": None, "토지면적_평": None,
                "신축총_평당_만원": None, "토지하한_평당_만원": None, "건물_사용연수": -1, "건물_구조": None,
                "추정방식": "오류(19자리 아님)", "근거지역": "", "참고": ""}

    지번 = jibun_from_pnu(p, dong_names)
    ic = lookup_idx(pnu_sorted, p)
    면적 = round(float(area_sorted[ic]), 1) if ic >= 0 and not np.isnan(area_sorted[ic]) else None
    면적평 = round(면적 / PYEONG, 1) if 면적 is not None else None
    det_age = int(age_sorted[ic]) if ic >= 0 else -1
    det_struct = {1: "rc", 2: "brick"}.get(int(struct_sorted[ic]) if ic >= 0 else 0)
    det_uq = bool(uq_sorted[ic]) if ic >= 0 else False

    def out(pair, 방식, 지역, 참고):
        T0, TL = float(pair[0]), float(pair[1])
        return {"PNU": p, "지번": 지번, "토지면적_㎡": 면적, "토지면적_평": 면적평,
                "신축총_평당_만원": round(T0, 1), "토지하한_평당_만원": round(TL, 1),
                "건물_사용연수": det_age, "건물_구조": det_struct, "지구단위": det_uq,
                "추정방식": 방식, "근거지역": 지역, "참고": 참고}

    # ① 실측 (실거래 [신축총, 토지하한] 중앙값)
    if p in parcel:
        return out(parcel[p], "실측", dong_names.get(p[:10], ""), "실거래")
    # ② 공간보간(동내) — 법정동 경계 안의 실거래만(하천·대로 미교차)
    if ic >= 0 and knn is not None:
        r = knn.idw_in_dong(xy_sorted[ic], p[:10])
        if r is not None:
            pair, nd = r
            return out(pair, "공간보간(동내)", "동내 실거래", f"최근접 {nd:.0f}m")
    # ③ 동 추정
    if p[:10] in dong:
        rec = dong[p[:10]]
        return out(rec[:2], "동 추정", dong_names.get(p[:10], ""), f"동 {rec[2] if len(rec)>2 else 0}지번")
    # ④ 구 추정
    if p[:5] in gu:
        rec = gu[p[:5]]
        return out(rec[:2], "구 추정", gu_names.get(p[:5], ""), f"구 {rec[2] if len(rec)>2 else 0}지번")
    # ⑤ 전체
    return out(total, "전체 추정", "서울 전체", "")


def residual_rate(age, structure, b):
    """국세청 공식 정액 잔가율 = max(최종잔존, 1 - 연상각률×사용연수)"""
    grp = b.get("brick" if structure == "brick" else "rc", {"rate": 0.018})
    return max(b.get("floor", 0.10), 1.0 - grp["rate"] * max(0, age))


def building_value(age, area, structure, b):
    """건물값(만원) = 건물원가(beta) × 잔가율(사용연수) × 건물면적(㎡)"""
    if area is None:
        return None
    return b.get("beta", 500.0) * residual_rate(age, structure, b) * area


def units_for(ho, pnu, base_pyeong):
    """호별 행 [호명, 층, 보정, 대지권㎡/평, 호평당가, 예상가]. base_pyeong=대지평당가(N년)"""
    rows = []
    for 호명, fc, 지분 in ho.get(pnu, []):
        m = FLOOR_MULT.get(fc, 1.0)
        호평당 = base_pyeong * m
        대지평 = 지분 / PYEONG
        rows.append({"호": 호명 or "-", "층": FLOOR_LABEL.get(fc, fc), "보정": m,
                     "대지권_㎡": round(지분, 2), "대지권_평": round(대지평, 2),
                     "호_대지평당가_만원": round(호평당, 1), "호_예상가_만원": round(호평당 * 대지평)})
    rows.sort(key=lambda r: (FLOOR_ORDER.get(r["층"], 9), r["호"]))
    return rows


# ====================================================================
#  GUI
# ====================================================================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("RE_price — 서울 빌라(연립다세대) 예상 거래가 산출기 (경량 배포판)")
        self.geometry("760x780"); self.minsize(700, 700)

        self.pnu_one = tk.StringVar()
        self.addr_one = tk.StringVar()
        self.age_one = tk.StringVar()
        self.bldarea_one = tk.StringVar()
        self.struct_one = tk.StringVar(value="rc")
        self.status = tk.StringVar(value="내장 데이터 준비 중 ... (잠시만 기다려 주세요)")
        self.log_q = queue.Queue()
        self.ui_q = queue.Queue()
        self.ctx = None
        self.ho = {}
        self.road = {}
        self.dong_rev = {}
        self.last_batch = None

        head = ttk.Frame(self); head.pack(fill="x", padx=10, pady=(8, 4))
        ttk.Label(head, text="빌라(연립다세대) 예상 거래가 산출 프로그램",
                  font=("", 12, "bold")).pack(anchor="w")
        desc = ("· 실거래가 = 토지+건물 합산. 대지평당가도 토지+건물 합산으로 산출하며 사용연수에 따라 감가.\n"
                "· 대지평당가(사용 N년) = 토지하한 + (신축총 − 토지하한) × 국세청 공식 잔가율(N년).\n"
                "  (철근콘크리트 50년·연1.8% / 벽돌·연와 40년·연2.25%, 최종잔존 10% 정액 — 건물만 감가)\n"
                "· 예상 거래가 = 대지평당가(N년) × 필지 평수. 오래될수록 토지하한으로 수렴(저가 매입 판단).")
        ttk.Label(head, text=desc, justify="left", foreground="#444").pack(anchor="w", pady=(2, 0))
        ttk.Label(head, textvariable=self.status, foreground="#2a6").pack(anchor="w", pady=(4, 0))

        bld = ttk.LabelFrame(self, text="감가 조건 (비우면 건축물대장에서 사용연수·구조 자동 적용 / 값 입력 시 수동 우선)")
        bld.pack(fill="x", padx=10, pady=(4, 0))
        ttk.Label(bld, text="사용연수").grid(row=0, column=0, padx=(8, 2), pady=6)
        ttk.Entry(bld, textvariable=self.age_one, width=6).grid(row=0, column=1)
        ttk.Label(bld, text="년 (비우면 자동)   구조(수동 시)").grid(row=0, column=2, padx=(10, 2))
        ttk.Radiobutton(bld, text="철근콘크리트", value="rc", variable=self.struct_one).grid(row=0, column=3, padx=2)
        ttk.Radiobutton(bld, text="벽돌·연와", value="brick", variable=self.struct_one).grid(row=0, column=4, padx=2)

        addr = ttk.LabelFrame(self, text="① 주소로 조회 (지번 또는 도로명)")
        addr.pack(fill="x", padx=10, pady=4)
        ttk.Entry(addr, textvariable=self.addr_one, width=46).grid(row=0, column=0, padx=6, pady=6)
        ttk.Button(addr, text="조회", command=self.query_addr).grid(row=0, column=1, padx=4)
        ttk.Label(addr, text="예: 강남구 역삼동 601-1  또는  영등포구 문래로 191", foreground="#999").grid(row=0, column=2, padx=4)

        many = ttk.LabelFrame(self, text="② 여러 개 조회 — 주소(지번/도로명) 목록을 줄마다 하나씩 붙여넣기")
        many.pack(fill="both", expand=False, padx=10, pady=4)
        self.paste = tk.Text(many, height=6, wrap="none")
        self.paste.pack(fill="x", padx=6, pady=6)
        btns = ttk.Frame(many); btns.pack(fill="x", padx=6, pady=(0, 6))
        ttk.Button(btns, text="조회", command=self.query_paste).pack(side="left")
        ttk.Button(btns, text="엑셀로 저장", command=self.save_excel).pack(side="left", padx=6)
        ttk.Button(btns, text="지우기", command=lambda: self.paste.delete("1.0", "end")).pack(side="left")

        ttk.Label(self, text="결과 / 진행 상황").pack(anchor="w", padx=10)
        self.log = tk.Text(self, height=12, wrap="word")
        self.log.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.after(100, self.drain)
        threading.Thread(target=self._prepare, daemon=True).start()

    def write(self, m): self.log.insert("end", m + "\n"); self.log.see("end")

    def drain(self):
        try:
            while True: self.write(self.log_q.get_nowait())
        except queue.Empty: pass
        try:
            while True:
                kind, val = self.ui_q.get_nowait()
                if kind == "status": self.status.set(val)
                elif kind == "ask": self._ask_folder()
        except queue.Empty: pass
        self.after(100, self.drain)

    def _prepare(self, extra=None):
        try:
            prices = load_prices(self.log_q.put, extra)
            pnu_s, xy_s, area_s, age_s, struct_s, uq_s = load_coords(self.log_q.put, extra)
            dn, gun = load_dongnames(self.log_q.put, extra)
            self.dong_rev = {name: code for code, name in dn.items()}
            knn = build_spatial(prices["parcel"], pnu_s, xy_s, self.log_q.put)
            ho = load_ho(self.log_q.put, extra)
            self.ho = ho
            self.road = load_road(self.log_q.put, extra)
            self.ctx = (prices["parcel"], prices["dong"], prices["gu"], prices["total_med"],
                        pnu_s, xy_s, area_s, dn, gun, knn,
                        prices.get("building", {"beta": 500.0, "floor": 0.10,
                                                "rc": {"rate": 0.018, "life": 50},
                                                "brick": {"rate": 0.0225, "life": 40}}),
                        age_s, struct_s, uq_s)
            self.ui_q.put(("status", f"✅ 준비 완료 — 실측 {len(prices['parcel']):,}지번 · 호데이터 {len(ho):,}필지. 조회하세요."))
            self.log_q.put("✅ 자동 반영 완료! 공간보간이 켜졌습니다.")
        except FileNotFoundError as e:
            self.ui_q.put(("status", "❌ 데이터를 찾지 못했습니다 — 폴더를 지정해 주세요"))
            self.log_q.put(f"[오류] {e}")
            self.ui_q.put(("ask", None))
        except Exception as e:
            self.ui_q.put(("status", "❌ 데이터 준비 실패"))
            self.log_q.put(f"[오류] {e}")

    def _ask_folder(self):
        messagebox.showinfo("데이터 폴더 지정",
                            "데이터를 자동으로 못 찾았습니다.\n"
                            "prices.json, PNU_coords.npz 가 있는 폴더를 골라주세요.")
        folder = filedialog.askdirectory(title="데이터 폴더 선택")
        if folder:
            self.status.set("선택한 폴더로 다시 준비 중 ...")
            threading.Thread(target=self._prepare, args=(folder,), daemon=True).start()

    def _won(self, manwon):
        if manwon is None:
            return "-"
        eok = manwon / 10000.0
        if eok >= 1:
            return f"{manwon:,.0f}만원 ({eok:,.1f}억)"
        return f"{manwon:,.0f}만원"

    def _manual_age(self):
        s = self.age_one.get().strip()
        if not s:
            return None
        try:
            return max(0, min(60, int(float(s))))
        except ValueError:
            return None

    def _bld_inputs(self):
        a = self._manual_age()
        return (a if a is not None else 0), self.struct_one.get()

    def _resolve(self, r):
        """사용연수·구조 결정: 수동입력 우선 → 건축물대장 자동 → 없음."""
        m = self._manual_age()
        if m is not None:
            return m, self.struct_one.get(), "수동"
        if r.get("건물_사용연수", -1) >= 0:
            return int(r["건물_사용연수"]), (r.get("건물_구조") or "rc"), "건축물대장"
        return None, None, None

    def _fmt(self, r, prefix="· "):
        if r.get("신축총_평당_만원") is None:
            return f"{prefix}{r['PNU']} → {r['추정방식']}"
        지번 = r.get("지번", "") or r["PNU"]
        면적 = (f"{r['토지면적_㎡']:,.1f}㎡ ({r['토지면적_평']:,.1f}평)"
              if r.get("토지면적_㎡") is not None else "면적정보없음")
        b = self.ctx[10] if self.ctx else {}
        T0 = r["신축총_평당_만원"]; TL = r["토지하한_평당_만원"]
        age, struct, src = self._resolve(r)
        has_ho = r["PNU"] in self.ho

        L = [f"{prefix}{지번}  (PNU {r['PNU']})"]
        if age is None and not has_ho:
            # 건축물 정보 전혀 없음 → 예상평당가(신축 기준)만
            L.append(f"    ▣ 예상 평당가 (신축 기준) = {T0:,.0f} 만원/평")
            L.append(f"    (건축물 정보 없음)")
            L.append(f"    필지면적 {면적} · [{r['추정방식']}{(', '+r['참고']) if r['참고'] else ''}]")
            return "\n".join(L)

        a = age if age is not None else 0
        st = struct or "rc"
        평당 = TL + (T0 - TL) * residual_rate(a, st, b)
        L.append(f"    ▣ 대지평당가 = {평당:,.0f} 만원/평")
        if age is not None:
            구조명 = "벽돌·연와" if st == "brick" else "철근콘크리트"
            태그 = "자동" if src == "건축물대장" else "수동"
            L.append(f"    건축물: {구조명} · 사용 {a}년 ({태그})")
        L.append(f"    신축 0년 상한 {T0:,.0f} 만원/평")
        rows = units_for(self.ho, r["PNU"], 평당)
        if rows:
            L.append("")
            L.append(f"    ▣ 호별 예상 ({len(rows)}호)")
            for u in rows[:60]:
                L.append(f"      · {u['호']}호 대지권 {u['대지권_㎡']}㎡({u['대지권_평']}평) · "
                         f"평당 {u['호_대지평당가_만원']:,.0f}만 · 예상가 {self._won(u['호_예상가_만원'])}")
            if len(rows) > 60:
                L.append(f"      ... 외 {len(rows)-60}호 ('엑셀로 저장'으로 전체 확인)")
        L.append("")
        L.append(f"    필지면적 {면적} · [{r['추정방식']}{(', '+r['참고']) if r['참고'] else ''}]")
        return "\n".join(L)

    def query_one(self):
        if not self.ctx: messagebox.showinfo("안내", "데이터 준비 중입니다."); return
        v = self.pnu_one.get().strip()
        if not v: messagebox.showinfo("안내", "PNU를 입력하세요."); return
        self.write(self._fmt(estimate_one(v, self.ctx)))

    def query_addr(self):
        if not self.ctx: messagebox.showinfo("안내", "데이터 준비 중입니다."); return
        a = self.addr_one.get().strip()
        if not a: messagebox.showinfo("안내", "지번 또는 도로명 주소를 입력하세요."); return
        pnu = resolve_address(a, self.dong_rev, self.road)
        if not pnu:
            self.write(f"· '{a}' → 주소를 PNU로 못 바꿨습니다. '자치구 동 번지' 형식인지 확인하세요.")
            return
        self.write(self._fmt(estimate_one(pnu, self.ctx), prefix=f"· [{a}] → "))

    def _parse(self):
        items = []
        for line in self.paste.get("1.0", "end").splitlines():
            s = line.strip().strip(",")
            if not s:
                continue
            digits = normalize_pnu(s)
            if len(digits) == 19:
                items.append(digits)
            else:
                p = resolve_address(s, self.dong_rev, self.road)
                if p:
                    items.append(p)
        return items

    def query_paste(self):
        if not self.ctx: messagebox.showinfo("안내", "데이터 준비 중입니다."); return
        pnus = self._parse()
        if not pnus: messagebox.showinfo("안내", "주소(지번/도로명)를 줄마다 하나씩 붙여넣으세요."); return
        threading.Thread(target=self._paste_work, args=(pnus,), daemon=True).start()

    def _paste_work(self, pnus):
        try:
            self.log_q.put(f"붙여넣은 {len(pnus):,}개 조회 중 ...")
            b = self.ctx[10] if self.ctx else {}
            m_age = self._manual_age()
            rows = []
            for p in pnus:
                r = estimate_one(p, self.ctx)
                if r.get("신축총_평당_만원") is not None:
                    T0 = r["신축총_평당_만원"]; TL = r["토지하한_평당_만원"]
                    # 수동입력 우선 → 건축물대장 자동 → 신축(0년)
                    if m_age is not None:
                        a, stc = m_age, self.struct_one.get()
                    elif r.get("건물_사용연수", -1) >= 0:
                        a, stc = int(r["건물_사용연수"]), (r.get("건물_구조") or "rc")
                    else:
                        a, stc = 0, "rc"
                    평당 = round(TL + (T0 - TL) * residual_rate(a, stc, b), 1)
                    면적평 = r.get("토지면적_평")
                    r["사용연수"] = a
                    r["구조"] = "벽돌·연와" if stc == "brick" else "철근콘크리트"
                    r["대지평당가_만원"] = 평당
                    r["예상거래가_만원"] = round(평당 * 면적평) if 면적평 is not None else None
                rows.append(r)
            self.last_batch = rows
            for r in rows[:30]:
                self.log_q.put(self._fmt(r))
            if len(rows) > 30:
                self.log_q.put(f"   ... 외 {len(rows)-30:,}건 ('엑셀로 저장'으로 전체 확인)")
            self.log_q.put(f"✅ {len(rows):,}건 완료(사용연수 자동). [엑셀로 저장]을 누르면 .xlsx로 저장됩니다.")
        except Exception as e:
            self.log_q.put(f"[오류] {e}")

    def save_excel(self):
        if not self.last_batch:
            messagebox.showinfo("안내", "먼저 ③에서 PNU를 붙여넣고 [조회]를 누르세요."); return
        path = filedialog.asksaveasfilename(
            title="엑셀로 저장", defaultextension=".xlsx",
            initialfile="RE_price_result.xlsx", filetypes=[("Excel 통합문서", "*.xlsx")])
        if not path: return
        try:
            wb = Workbook()
            # === 시트1: 조회 PNU 평당가 리스트 ===
            cols = ["PNU", "지번", "토지면적_㎡", "토지면적_평", "사용연수", "구조",
                    "대지평당가_만원", "예상거래가_만원", "신축총_평당_만원", "토지하한_평당_만원", "추정방식", "참고"]
            ws1 = wb.active; ws1.title = "평당가"
            ws1.append(cols)
            for r in self.last_batch:
                ws1.append([r.get(c) for c in cols])
            # === 시트2: 호별 데이터 (호 정보 있는 지번만) ===
            ws2 = wb.create_sheet("호별")
            hcols = ["지번", "PNU", "호", "층", "대지권_㎡", "대지권_평",
                     "호_대지평당가_만원", "호_예상가_만원", "사용연수", "구조"]
            ws2.append(hcols)
            n_ho = 0; n_parcel = 0
            for r in self.last_batch:
                base = r.get("대지평당가_만원")
                if base is None or r["PNU"] not in self.ho:
                    continue
                n_parcel += 1
                for u in units_for(self.ho, r["PNU"], base):
                    ws2.append([r.get("지번"), r["PNU"], u["호"], u["층"], u["대지권_㎡"],
                                u["대지권_평"], u["호_대지평당가_만원"], u["호_예상가_만원"],
                                r.get("사용연수"), r.get("구조")])
                    n_ho += 1
            wb.save(path)
            self.write(f"✅ 엑셀 저장 완료 → {path}\n    시트1 평당가 {len(self.last_batch):,}건 · 시트2 호별 {n_ho:,}호({n_parcel:,}지번)")
            messagebox.showinfo("완료", f"시트1 평당가 {len(self.last_batch):,}건\n시트2 호별 {n_ho:,}호 ({n_parcel:,}지번)\n{path}")
        except Exception as e:
            messagebox.showerror("오류", f"엑셀 저장 실패: {e}")


if __name__ == "__main__":
    App().mainloop()

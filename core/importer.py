"""
Import file Excel "Usulan Limit Counterparty" (format seperti Limit_Countreparty_09102026.xlsx)
ke master data historis. Dipakai untuk:
  1) membangun master data awal, dan
  2) memasukkan file-file periode lama sehingga riwayat rating terbentuk.

Pencarian posisi memakai label (bukan alamat sel tetap), jadi file dengan jumlah
baris counterparty berbeda tetap bisa dibaca.
"""
from __future__ import annotations

import re
from datetime import date, datetime
from io import BytesIO

import pandas as pd
from openpyxl import load_workbook

from .defaults import default_untuk
from .engine import (JENIS_EKSTERNAL, JENIS_INTERNAL, KAT_MI, KAT_NON_MI, L_INTERNAL, L_PEFINDO, P_MI_ARAHAN,
                     P_MI_D, P_NONMI_A, STATUS_KHUSUS, STATUS_NONAKTIF, STATUS_OVERRIDE, STATUS_RUMUS, kosong,
                     lembaga_dari_rating, norm_key, normalize_master, parse_rating)


def _num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def _txt(v) -> str:
    return "" if v is None else re.sub(r"\s+", " ", str(v)).strip()


def _date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return None


def _tanggal_dari_nama_file(nama: str):
    m = re.search(r"(\d{2})(\d{2})(\d{4})", nama or "")
    if m:
        try:
            return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            return None
    return None


def parse_snapshot(data: bytes, nama_file: str = "") -> dict:
    """Baca satu file laporan limit (format Excel asli) menjadi snapshot terstruktur."""
    wf = load_workbook(BytesIO(data), data_only=False)
    wv = load_workbook(BytesIO(data), data_only=True)

    ws_f = ws_v = None
    for name in wf.sheetnames:
        s = wf[name]
        for row in s.iter_rows(min_col=2, max_col=2):
            if norm_key(row[0].value) == "TOTAL NON MANAJER INVESTASI":
                ws_f, ws_v = s, wv[name]
                break
        if ws_f:
            break
    if ws_f is None:
        raise ValueError("Tidak ditemukan baris 'TOTAL NON MANAJER INVESTASI' di kolom B. Format file tidak dikenali.")

    maxr, maxc = ws_f.max_row, ws_f.max_column
    F = lambda r, c: ws_f.cell(r, c).value
    V = lambda r, c: ws_v.cell(r, c).value

    # --- tanggal laporan: baris yang berisi "per tanggal"
    tgl = None
    for r in range(1, 6):
        for c in range(1, min(maxc, 12) + 1):
            if "PER TANGGAL" in norm_key(F(r, c)):
                for c2 in range(c + 1, min(maxc, 12) + 1):
                    tgl = _date(V(r, c2))
                    if tgl:
                        break
        if tgl:
            break
    tgl = tgl or _tanggal_dari_nama_file(nama_file)

    # --- lokasi baris kunci
    r_nonmi = r_mi = None
    for r in range(1, maxr + 1):
        k = norm_key(F(r, 2))
        if k == "TOTAL NON MANAJER INVESTASI":
            r_nonmi = r
        elif k == "TOTAL MANAJER INVESTASI":
            r_mi = r

    def baca_baris(r0, kat):
        rows, r = [], r0 + 1
        while r <= maxr and _num(F(r, 2)) is not None and _txt(F(r, 3)):
            d_f = F(r, 4)
            is_formula = isinstance(d_f, str) and d_f.startswith("=")
            nama = _txt(F(r, 3))
            tanda = _txt(F(r, 7)).lower()
            aktual = _num(V(r, 5))
            if is_formula:
                status, limit = STATUS_RUMUS, None
                if kat == KAT_NON_MI and tanda != "v":
                    status = STATUS_NONAKTIF
            else:
                limit = _num(d_f)
                if kat == KAT_NON_MI and tanda != "v":
                    status = STATUS_NONAKTIF
                elif "+)" in nama:
                    status = STATUS_KHUSUS
                else:
                    status = STATUS_OVERRIDE
            rows.append({"Nama": nama, "Kategori": kat, "Rating": _txt(F(r, 6)) or None,
                         "Status": status, "Limit": limit, "Aktual": aktual,
                         "Limit Excel": _num(V(r, 4)), "Baris": r, "Tanda": tanda})
            r += 1
        return rows, r

    nonmi, r_end_nonmi = baca_baris(r_nonmi, KAT_NON_MI)
    mi, _ = baca_baris(r_mi, KAT_MI) if r_mi else ([], None)

    # --- catatan kaki
    catatan = []
    batas = (r_mi - 2) if r_mi else min(maxr, r_end_nonmi + 6)
    for r in range(r_end_nonmi, batas + 1):
        t = _txt(F(r, 3))
        if t and len(t) > 15:
            catatan.append({"Bagian": KAT_NON_MI, "Teks": t})
    if r_mi:
        for r in range(r_mi, maxr + 1):
            t = _txt(F(r, 8))
            if t.startswith("+)"):
                catatan.append({"Bagian": KAT_MI, "Teks": t})

    # --- parameter blok Non MI (kolom H = 8, nilai mulai kolom I = 9)
    blok_end = (r_mi - 1) if r_mi else maxr
    total_aset = ta_tgl = a_nonmi = None
    tiers, bobot_ext = [], []
    for r in range(1, blok_end + 1):
        h = norm_key(F(r, 8))
        if total_aset is None and h.startswith("(D) TOTAL ASET INVESTASI"):
            total_aset = _num(V(r, 9))
            if total_aset is None:
                total_aset = _num(V(r - 1, 9))
            ta_tgl = _date(V(r - 1, 10)) or tgl
        elif a_nonmi is None and h.startswith("(A)") and _num(V(r, 9)) is not None:
            a_nonmi = _num(V(r, 9))
        elif not bobot_ext and h.startswith("(B)") and _num(V(r, 9)) is not None:
            c = 9
            while c <= maxc and _num(V(r, c)) is not None:
                bobot_ext.append(_num(V(r, c)))
                c += 1
    # label tier: baris di kolom I berisi "idXXX" (tanpa akhiran)
    for r in range(1, min(blok_end, 15) + 1):
        lab = _txt(F(r, 9))
        p = parse_rating(lab)
        if p and p["id"] and not p["extra"] and not p["idn"]:
            c = 9
            while c <= maxc and parse_rating(_txt(F(r, c))):
                tiers.append(parse_rating(_txt(F(r, c)))["core"])
                c += 1
            break

    # --- parameter blok MI
    mi_arahan = mi_d = None
    int_labels, bobot_int = [], []
    if r_mi:
        for r in range(r_mi, maxr + 1):
            h = norm_key(F(r, 8))
            if not int_labels and h.startswith("RATING INTERNAL"):
                c = 9
                while c <= maxc and _txt(F(r, c)):
                    int_labels.append(_txt(F(r, c)))
                    c += 1
            elif mi_arahan is None and h.startswith("(A)") and "ARAHAN" in h:
                mi_arahan = _num(V(r, 9))
            elif mi_d is None and h == "(D)" and _num(V(r, 9)) is not None:
                mi_d = _num(V(r, 9))
            elif not bobot_int and h == "(E)" and _num(V(r, 9)) is not None:
                c = 9
                while c <= maxc and _num(V(r, c)) is not None:
                    bobot_int.append(_num(V(r, c)))
                    c += 1

    warn = []
    if not tgl:
        warn.append("Tanggal laporan tidak terbaca — isi manual.")
    if total_aset is None:
        warn.append("Total Aset Investasi tidak terbaca.")
    if len(tiers) != len(bobot_ext):
        warn.append("Jumlah label rating eksternal ≠ jumlah bobot; bobot eksternal tidak diimpor.")
        tiers, bobot_ext = [], []
    if len(int_labels) != len(bobot_int):
        warn.append("Jumlah label rating internal ≠ jumlah bobot; bobot internal tidak diimpor.")
        int_labels, bobot_int = [], []

    return {
        "file": nama_file, "tanggal": tgl, "total_aset": total_aset, "total_aset_tgl": ta_tgl,
        "rows": nonmi + mi, "catatan": catatan,
        "params": {k: v for k, v in {P_NONMI_A: a_nonmi, P_MI_ARAHAN: mi_arahan, P_MI_D: mi_d}.items() if v is not None},
        "bobot": {JENIS_EKSTERNAL: list(zip(tiers, bobot_ext)), JENIS_INTERNAL: list(zip(int_labels, bobot_int))},
        "warnings": warn,
    }


# --------------------------------------------------------------------------
def _compress(df: pd.DataFrame, key_cols: list[str] | None, date_col: str, val_cols: list[str]) -> pd.DataFrame:
    """Urutkan per kunci & tanggal, lalu buang baris yang nilainya sama dengan baris sebelumnya
    (sehingga riwayat hanya berisi titik perubahan, dengan tanggal paling awal)."""
    if df.empty:
        return df
    d = df.copy()
    d["_i"] = range(len(d))
    d["_k"] = d[key_cols].apply(lambda r: "|".join(norm_key(v) for v in r), axis=1) if key_cols else ""
    d = d.sort_values(["_k", date_col, "_i"], kind="stable")
    d = d.groupby(["_k", date_col], sort=False, dropna=False).tail(1)
    sig = d[val_cols].apply(lambda r: "|".join("" if kosong(v) else str(v) for v in r), axis=1)
    prev_sig = sig.groupby(d["_k"]).shift(1)
    d = d[sig != prev_sig]
    return d.drop(columns=["_i", "_k"]).reset_index(drop=True)


def gabung_rating(df: pd.DataFrame) -> pd.DataFrame:
    """Gabungkan baris rating berurutan yang nilainya sama (per Nama+Lembaga) menjadi satu rentang:
    Tanggal Efektif = awal rentang, Terakhir Dicek = konfirmasi terakhir, Outlook = yang terbaru."""
    if df.empty:
        return df
    d = df.copy()
    d["_i"] = range(len(d))
    d["_k"] = d["Nama"].map(norm_key) + "|" + d["Lembaga"].map(norm_key)
    d = d.sort_values(["_k", "Tanggal Efektif", "_i"], kind="stable")
    # baris ganda (kunci & tanggal sama): baris terakhir menang, tetapi tanggal konfirmasi & outlook tidak hilang
    grp = d.groupby(["_k", "Tanggal Efektif"], sort=False, dropna=False)
    d["Terakhir Dicek"] = grp["Terakhir Dicek"].transform("max")
    d["Outlook"] = grp["Outlook"].transform(lambda s: s.dropna().iloc[-1] if s.notna().any() else None)
    d = d.groupby(["_k", "Tanggal Efektif"], sort=False, dropna=False).tail(1)
    sig = d["Rating"].map(norm_key)
    run = ((sig != sig.groupby(d["_k"]).shift(1)) | (d["_k"] != d["_k"].shift(1))).cumsum()
    out = []
    for _, g in d.groupby(run, sort=False):
        first = g.iloc[0].copy()
        akhir = pd.concat([g["Tanggal Efektif"], g["Terakhir Dicek"]]).dropna().max()
        first["Terakhir Dicek"] = akhir if akhir > first["Tanggal Efektif"] else first["Terakhir Dicek"]
        ol = g["Outlook"].dropna()
        if not ol.empty:
            first["Outlook"] = ol.iloc[-1]
        out.append(first)
    res = pd.DataFrame(out).drop(columns=["_i", "_k"])
    return res.reset_index(drop=True)


def merge_snapshot(master: dict, snap: dict, tanggal: date | None = None) -> tuple[dict, dict]:
    m = {k: v.copy() for k, v in master.items()}
    tgl = pd.Timestamp(tanggal or snap["tanggal"])
    if pd.isna(tgl):
        raise ValueError("Tanggal laporan wajib diisi.")
    sumber = f"Import {snap.get('file') or 'Excel'}".strip()
    ring = {"counterparty_baru": 0, "rating": 0, "status_limit": 0}

    # Counterparty
    cp = m["Counterparty"]
    known = set(cp["Nama"].map(norm_key))
    no_per_kat = {k: int(g["No"].max()) for k, g in cp.groupby("Kategori")} if not cp.empty else {}
    add = []
    for r in snap["rows"]:
        if norm_key(r["Nama"]) in known:
            continue
        kat = r["Kategori"]
        no_per_kat[kat] = no_per_kat.get(kat, 0) + 1
        if kat == KAT_MI:
            kode = entitas = alias = ket = ""
            lem = L_INTERNAL
        else:
            kode, entitas, alias, ket = default_untuk(r["Nama"])
            lem = lembaga_dari_rating(r["Rating"]) or L_PEFINDO
        add.append({"No": no_per_kat[kat], "Nama": r["Nama"], "Kategori": kat, "Lembaga Utama": lem,
                    "Kode PEFINDO": kode or None, "Nama Entitas": entitas or None, "Alias": alias or None,
                    "Keterangan": ket or None})
        known.add(norm_key(r["Nama"]))
    if add:
        m["Counterparty"] = pd.concat([cp, pd.DataFrame(add)], ignore_index=True)
    ring["counterparty_baru"] = len(add)

    # Rating
    n0 = len(m["Rating"])
    rt = pd.DataFrame([{
        "Nama": r["Nama"], "Tanggal Efektif": tgl, "Rating": r["Rating"],
        "Lembaga": L_INTERNAL if r["Kategori"] == KAT_MI else (lembaga_dari_rating(r["Rating"]) or None),
        "Terakhir Dicek": tgl, "Sumber": sumber,
    } for r in snap["rows"] if r["Rating"]])
    if not rt.empty:
        base = normalize_master({**m, "Rating": pd.concat([m["Rating"], rt], ignore_index=True)})["Rating"]
        m["Rating"] = gabung_rating(base)
    ring["rating"] = len(m["Rating"]) - n0

    # Status limit (simpan state lengkap lalu kompres; buang "Rumus" di awal riwayat)
    n0 = len(m["Status_Limit"])
    ov = pd.DataFrame([{
        "Nama": r["Nama"], "Tanggal Efektif": tgl, "Limit Nominal": r["Limit"], "Status": r["Status"],
        "Keterangan": sumber,
    } for r in snap["rows"]])
    if not ov.empty:
        d = _compress(pd.concat([m["Status_Limit"], ov], ignore_index=True), ["Nama"], "Tanggal Efektif",
                      ["Limit Nominal", "Status"])
        d["_k"] = d["Nama"].map(norm_key)
        first = ~d["_k"].duplicated()
        d = d[~(first & (d["Status"] == STATUS_RUMUS))].drop(columns="_k").reset_index(drop=True)
        m["Status_Limit"] = d
    ring["status_limit"] = len(m["Status_Limit"]) - n0

    # Total aset (upsert per tanggal)
    if snap["total_aset"] is not None:
        tt = pd.Timestamp(snap["total_aset_tgl"] or tgl)
        d = m["Total_Aset"]
        d = d[d["Tanggal"] != tt]
        m["Total_Aset"] = pd.concat([d, pd.DataFrame([{"Tanggal": tt, "Total Aset Investasi": snap["total_aset"],
                                                         "Keterangan": sumber}])], ignore_index=True
                                    ).sort_values("Tanggal", kind="stable").reset_index(drop=True)

    # Parameter
    if snap["params"]:
        pr = pd.DataFrame([{"Tanggal Efektif": tgl, "Kode": k, "Nilai": v, "Keterangan": sumber}
                           for k, v in snap["params"].items()])
        m["Parameter"] = _compress(pd.concat([m["Parameter"], pr], ignore_index=True), ["Kode"], "Tanggal Efektif",
                                   ["Nilai"])

    # Bobot rating: tambahkan set baru hanya jika berbeda dari set yang berlaku pada tanggal tsb
    bw = m["Bobot_Rating"]
    for jenis, pairs in snap["bobot"].items():
        if not pairs:
            continue
        cur = bw[(bw["Jenis"] == jenis) & (bw["Tanggal Efektif"] <= tgl)]
        same = False
        if not cur.empty:
            t0 = cur["Tanggal Efektif"].max()
            cur = cur[cur["Tanggal Efektif"] == t0].sort_values("Urutan")
            same = [(norm_key(a), round(float(b), 10)) for a, b in zip(cur["Rating"], cur["Bobot"])] == \
                   [(norm_key(a), round(float(b), 10)) for a, b in pairs]
        if not same:
            bw = pd.concat([bw, pd.DataFrame([{"Tanggal Efektif": tgl, "Jenis": jenis, "Urutan": i + 1,
                                                "Rating": a, "Bobot": b} for i, (a, b) in enumerate(pairs)])],
                           ignore_index=True)
    m["Bobot_Rating"] = bw

    # Eksposur (jika kolom Aktual Nominal terisi)
    ek_new = [{"Tanggal": tgl, "Nama": r["Nama"], "Aktual Nominal": r["Aktual"]} for r in snap["rows"] if r["Aktual"] is not None]
    if ek_new:
        d = m["Eksposur"]
        d = d[d["Tanggal"] != tgl]
        m["Eksposur"] = pd.concat([d, pd.DataFrame(ek_new)], ignore_index=True)

    # Catatan
    if snap["catatan"]:
        ada = set(m["Catatan"]["Teks"].map(norm_key))
        baru = [c for c in snap["catatan"] if norm_key(c["Teks"]) not in ada]
        if baru:
            m["Catatan"] = pd.concat([m["Catatan"], pd.DataFrame(baru)], ignore_index=True)

    return normalize_master(m), ring

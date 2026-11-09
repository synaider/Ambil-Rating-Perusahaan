"""
Baca file Excel limit (format asli) dan tulis ulang dengan nilai periode baru.

Yang diubah hanya:
  * tanggal laporan ("Usulan Limit ... per tanggal :")
  * rating di kolom Rating (counterparty Non Manajer Investasi)
  * Total Aset Investasi & tanggal posisinya (jika diisi)
  * (opsional) rentang rumus SUM total agar mencakup semua baris
Format, rumus limit, catatan, dan bagian Manajer Investasi tidak diubah.
"""
from __future__ import annotations

import re
from datetime import date, datetime
from io import BytesIO

from openpyxl import load_workbook

from .engine import KAT_NON_MI, P_NONMI_A, excel_round, norm_key, tier_eksternal
from .importer import parse_snapshot

RE_REF = re.compile(r"^=\s*\$?([A-Z]{1,3})\$?(\d+)\s*$")
RE_SUM = re.compile(r"^=\s*SUM\(\s*([A-Z]{1,3})(\d+)\s*:\s*([A-Z]{1,3})(\d+)\s*\)\s*$", re.I)


def baca_template(data: bytes, nama_file: str = "") -> dict:
    snap = parse_snapshot(data, nama_file)
    snap["nonmi"] = [r for r in snap["rows"] if r["Kategori"] == KAT_NON_MI]
    return snap


def hitung_limit(snap: dict, rating: dict[str, str | None], total_aset: float | None = None) -> dict[str, float | None]:
    """Limit Nominal per counterparty Non MI dengan rumus yang sama seperti di file.
    Baris yang di file berisi angka tetap (bukan rumus) memakai angka itu."""
    ta = total_aset if total_aset is not None else snap["total_aset"]
    a = snap["params"].get(P_NONMI_A)
    tiers = snap["bobot"].get("Eksternal") or []
    tlist = [t for t, _ in tiers]
    bobot = dict(tiers)
    out = {}
    for r in snap["nonmi"]:
        if r["Limit"] is not None:                       # angka tetap di file
            out[r["Nama"]] = float(r["Limit"])
            continue
        t = tier_eksternal(rating.get(r["Nama"]), tlist)
        out[r["Nama"]] = excel_round((a * bobot[t]) * ta, 0) if (t and a is not None and ta is not None) else None
    return out


def _sheet_utama(wb):
    for ws in wb.worksheets:
        for row in ws.iter_rows(min_col=2, max_col=2):
            if norm_key(row[0].value) == "TOTAL NON MANAJER INVESTASI":
                return ws
    raise ValueError("Format file tidak dikenali (baris 'TOTAL NON MANAJER INVESTASI' tidak ada).")


def tulis_template(data: bytes, snap: dict, tanggal: date, rating_baru: dict[str, str | None],
                   total_aset: float | None = None, tgl_aset: date | None = None,
                   perbaiki_total: bool = True) -> bytes:
    wb = load_workbook(BytesIO(data))
    ws = _sheet_utama(wb)
    maxc = min(ws.max_column, 20)

    # 1) tanggal laporan
    ditulis = False
    for r in range(1, 6):
        for c in range(1, maxc + 1):
            if "PER TANGGAL" in norm_key(ws.cell(r, c).value):
                for c2 in range(c + 1, maxc + 1):
                    if isinstance(ws.cell(r, c2).value, (datetime, date)):
                        ws.cell(r, c2).value = datetime(tanggal.year, tanggal.month, tanggal.day)
                        ditulis = True
                        break
                if not ditulis:
                    cell = ws.cell(r, max(c + 3, 5))
                    cell.value = datetime(tanggal.year, tanggal.month, tanggal.day)
                    cell.number_format = "dd mmmm yyyy"
                    ditulis = True
                break
        if ditulis:
            break

    # 2) rating Non MI (kolom F)
    for r in snap["nonmi"]:
        baru = rating_baru.get(r["Nama"])
        if baru:
            ws.cell(r["Baris"], 6).value = baru

    # 3) Total Aset Investasi
    if total_aset is not None:
        for rr in range(1, ws.max_row + 1):
            if norm_key(ws.cell(rr, 8).value).startswith("(D) TOTAL ASET INVESTASI"):
                target = ws.cell(rr, 9)
                m = RE_REF.match(str(target.value or ""))
                if m:
                    target = ws[f"{m.group(1)}{m.group(2)}"]
                target.value = float(total_aset)
                if tgl_aset:
                    tc = ws.cell(target.row, target.column + 1)
                    if tc.value is None or isinstance(tc.value, (datetime, date)):
                        tc.value = datetime(tgl_aset.year, tgl_aset.month, tgl_aset.day)
                        tc.number_format = "dd/mm/yyyy"
                break

    # 4) rumus total mencakup semua baris
    if perbaiki_total:
        for kat, judul in ((KAT_NON_MI, "TOTAL NON MANAJER INVESTASI"), ("MI", "TOTAL MANAJER INVESTASI")):
            baris = [r["Baris"] for r in snap["rows"] if r["Kategori"] == kat]
            if not baris:
                continue
            akhir = max(baris)
            for rr in range(1, ws.max_row + 1):
                if norm_key(ws.cell(rr, 2).value) == judul:
                    for c in (4, 5):
                        m = RE_SUM.match(str(ws.cell(rr, c).value or ""))
                        if m and int(m.group(4)) < akhir:
                            ws.cell(rr, c).value = f"=SUM({m.group(1)}{m.group(2)}:{m.group(3)}{akhir})"
                    break

    wb.calculation.fullCalcOnLoad = True
    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()

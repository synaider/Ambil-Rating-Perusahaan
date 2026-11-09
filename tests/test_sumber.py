"""Uji urutan pencarian rating (PEFINDO → Fitch → anggota group), filter tanggal, dan penulisan Excel."""
import sys
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from core.pefindo import DataEmiten  # noqa: E402
from core.sumber import KOLOM_FITCH, Temuan, cek_group, keterangan, rapikan_entitas, rapikan_fitch  # noqa: E402
from core.template_excel import baca_template, hitung_limit, tulis_template  # noqa: E402


def emiten(kode, rilis, rating=None, upd=None):
    return DataEmiten(kode=kode, url=f"https://www.pefindo.com/rating-action-reports/rating-report/{kode}",
                      rating=rating, outlook="Stable" if rating else None, tgl_update=upd, rilis=rilis)


class PefindoPalsu:
    def __init__(self, data):
        self.data = data

    def ambil(self, kode):
        return (self.data[kode], "") if kode in self.data else (None, f"kode {kode} tidak ada di pefindo.com")


ENT = rapikan_entitas(pd.DataFrame([
    {"Group": "SALIM GROUP", "Urutan": 1, "Entitas": "PT Indofood Sukses Makmur Tbk", "Kode PEFINDO": "INDF"},
    {"Group": "SALIM GROUP", "Urutan": 2, "Entitas": "PT Salim Ivomas Pratama Tbk", "Kode PEFINDO": "SIMP"},
    {"Group": "OCBC GROUP", "Urutan": 1, "Entitas": "PT Bank OCBC NISP Tbk", "Kode PEFINDO": "NISP"},
    {"Group": "DBS GROUP", "Urutan": 1, "Entitas": "PT Bank DBS Indonesia", "Kode PEFINDO": None},
    {"Group": "BRI GROUP", "Urutan": 1, "Entitas": "PT Bank Rakyat Indonesia (Persero) Tbk", "Kode PEFINDO": "BBRI"},
]))
FIT = rapikan_fitch(pd.DataFrame([
    {"Entitas": "PT Bank OCBC NISP Tbk", "Tanggal": "2026-03-12", "Rating": "AAA(idn)"},
    {"Entitas": "PT Bank DBS Indonesia", "Tanggal": "2025-12-01", "Rating": "AAA(idn)"},
    {"Entitas": "PT Bank DBS Indonesia", "Tanggal": "2026-11-05", "Rating": "AA+(idn)"},
], columns=KOLOM_FITCH))
PEF = PefindoPalsu({
    # Indofood: hanya rilis instrumen → tidak ada rating perusahaan
    "INDF": emiten("INDF", [(date(2026, 5, 1), "PEFINDO has affirmed its idAA+ rating for PT Indofood Sukses Makmur "
                                               "Tbk's Bond IX")]),
    "SIMP": emiten("SIMP", [(date(2026, 2, 10), "PEFINDO has affirmed its idAA rating to PT Salim Ivomas Pratama Tbk")],
                   "idAA", date(2026, 2, 10)),
    # OCBC: rating PEFINDO terakhir 2019 → sudah tidak berlaku → pakai Fitch
    "NISP": emiten("NISP", [(date(2019, 6, 1), 'PEFINDO has affirmed its "idAAA" rating to PT Bank OCBC NISP Tbk')]),
    # BRI: update 9 Nov 2026 (turun), sebelum itu idAAA
    "BBRI": emiten("BBRI", [(date(2025, 9, 19), "PEFINDO has affirmed its idAAA ratings to PT Bank Rakyat Indonesia "
                                                "(Persero) Tbk (BBRI) and its outstanding bonds"),
                            (date(2026, 9, 3), "PEFINDO has affirmed its idAAA ratings to PT Bank Rakyat Indonesia "
                                               "(Persero) Tbk (BBRI) and its outstanding bonds"),
                            (date(2026, 11, 9), "PEFINDO has lowered the corporate rating of PT Bank Rakyat Indonesia "
                                                "(Persero) Tbk (BBRI) to idAA+ from idAAA")],
                   "idAA+", date(2026, 11, 9)),
})


def test_fallback_anggota_group():
    tm = cek_group("SALIM GROUP", ENT, FIT, date(2026, 11, 9), PEF)
    assert (tm.rating, tm.lembaga, tm.urutan) == ("idAA", "PEFINDO", 2), tm
    assert "anggota group" in keterangan(tm)


def test_pefindo_kedaluwarsa_pakai_fitch():
    tm = cek_group("OCBC GROUP", ENT, FIT, date(2026, 11, 9), PEF)
    assert (tm.rating, tm.lembaga) == ("AAA(idn)", "Fitch"), tm
    assert "PEFINDO tidak ada" in keterangan(tm)


def test_tanggal_periode():
    # PEFINDO update 9/11 → cek 8/11 harus dapat rating sebelumnya
    assert cek_group("BRI GROUP", ENT, FIT, date(2026, 11, 8), PEF).rating == "idAAA"
    assert cek_group("BRI GROUP", ENT, FIT, date(2026, 11, 9), PEF).rating == "idAA+"
    # Fitch DBS: 5/11 berubah
    assert cek_group("DBS GROUP", ENT, FIT, date(2026, 11, 4), PEF).rating == "AAA(idn)"
    assert cek_group("DBS GROUP", ENT, FIT, date(2026, 11, 5), PEF).rating == "AA+(idn)"
    # sebelum ada data sama sekali
    assert cek_group("DBS GROUP", ENT, FIT, date(2025, 1, 1), PEF).rating is None


def test_tidak_ada_entitas():
    tm = cek_group("GROUP BARU", ENT, FIT, date(2026, 11, 9), PEF)
    assert tm.rating is None and "Tidak ditemukan" in keterangan(tm)


def test_tulis_template():
    f = ROOT / "data" / "contoh" / "Limit_Countreparty_09102026.xlsx"
    data = f.read_bytes()
    snap = baca_template(data, f.name)
    assert len(snap["nonmi"]) == 44
    baru = {r["Nama"]: r["Rating"] for r in snap["nonmi"]}
    baru["SALIM GROUP"] = "idAA"
    baru["OCBC Group"] = "AA+(idn)"
    lim = hitung_limit(snap, baru)
    assert lim["SALIM GROUP"] == 433 and lim["OCBC Group"] == 525 and lim["BRI GROUP+)"] == 725
    out = tulis_template(data, snap, date(2026, 11, 8), baru, total_aset=6500.5, tgl_aset=date(2026, 10, 31))
    ws = load_workbook(__import__("io").BytesIO(out)).active
    assert ws["E2"].value.date() == date(2026, 11, 8)
    assert ws["F28"].value == "idAA" and ws["F47"].value == "AA+(idn)"
    assert ws["I7"].value == 6500.5 and ws["J7"].value.date() == date(2026, 10, 31)
    assert ws["D6"].value == "=SUM(D7:D50)" and ws["E6"].value == "=SUM(E7:E50)"
    assert str(ws["D8"].value).startswith("=IF(")          # rumus asli tetap
    assert ws["D7"].value == 725                            # angka tetap BRI tidak diubah
    assert ws["F58"].value == "Sangat Bagus"                # bagian MI tidak diubah


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)


# --- pengaturan sumber yang ditetapkan pengguna (9 Nov 2026) -------------------------------
from core.pefindo import riwayat_dari_emiten  # noqa: E402
from core.sumber import entitas_default, rating_pefindo_pada, sesuaikan_versi_lama  # noqa: E402

ASMS = emiten("ASMS", [
    (date(2023, 1, 19), 'PEFINDO has assigned its "idAA+" financial strength rating to PT Asuransi Sinar Mas (ASM)'),
    (date(2024, 1, 30), 'PEFINDO has assigned its "idAA+" financial strength rating with "stable" outlook to PT Asuransi '
                        'Sinar Mas (ASM)'),
    (date(2025, 2, 6), "PEFINDO has assigned its idAA+ financial strength rating with a stable outlook to PT Asuransi "
                       "Sinar Mas (ASM)"),
    (date(2026, 2, 24), "PEFINDO has assigned its idAA+ financial strength rating with a stable outlook to PT Asuransi "
                        "Sinar Mas (ASM)")], "idAA+", date(2026, 2, 24))


def test_default_sumber_pengguna():
    d = entitas_default()
    sm = d[d["Group"] == "SINARMAS GROUP"]
    assert list(sm["Kode PEFINDO"]) == ["ASMS"]
    assert list(d[d["Group"] == "BARITO PACIFIC GROUP"]["Kode PEFINDO"]) == ["BRPT"]
    assert list(d[d["Group"] == "BANK BTPN GROUP"]["Lembaga"]) == ["Fitch"]


def test_asuransi_sinar_mas():
    rows, _ = riwayat_dari_emiten(ASMS)
    assert [r["Rating"] for r in rows] == ["idAA+"] * 4
    assert rows[1]["Outlook"] == "Stable"
    r, _ = rating_pefindo_pada(ASMS, date(2026, 11, 9))
    assert r["Rating"] == "idAA+" and r["Tanggal Efektif"].date() == date(2026, 2, 24)
    ent = entitas_default()
    tm = cek_group("SINARMAS GROUP", ent, FIT, date(2026, 11, 9), PefindoPalsu({"ASMS": ASMS}))
    assert (tm.rating, tm.lembaga, tm.entitas) == ("idAA+", "PEFINDO", "PT Asuransi Sinar Mas")


def test_btpn_dari_fitch():
    ent = entitas_default()
    fit = rapikan_fitch(pd.DataFrame([{"Entitas": "PT Bank SMBC Indonesia Tbk", "Tanggal": "2026-10-09",
                                       "Rating": "AAA(idn)"}], columns=KOLOM_FITCH))
    pef_btpn = PefindoPalsu({"BTPN": emiten("BTPN", [(date(2026, 6, 1), "PEFINDO has affirmed its idAA rating to "
                                                                        "PT Bank SMBC Indonesia Tbk")], "idAA",
                                            date(2026, 6, 1))})
    tm = cek_group("Bank BTPN Group", ent, fit, date(2026, 11, 9), pef_btpn)
    assert (tm.rating, tm.lembaga) == ("AAA(idn)", "Fitch"), tm      # PEFINDO dilewati
    assert "sumber diatur: Fitch" in keterangan(tm)
    assert "fitchratings.com" in tm.url
    # SMFG (Otomatis) tetap boleh memakai PEFINDO untuk entitas yang sama
    tm2 = cek_group("SUMITOMO MITSUI FINANCIAL GROUP", ent, fit, date(2026, 11, 9), pef_btpn)
    assert (tm2.rating, tm2.lembaga) == ("idAA", "PEFINDO")


def test_pengaturan_versi_lama_disesuaikan():
    lama = entitas_default().drop(columns="Lembaga")
    # versi lama: Sinarmas mulai dari SMAR, Barito punya 3 entitas, BTPN dari PEFINDO
    lama = lama[~lama["Group"].isin(["SINARMAS GROUP", "BARITO PACIFIC GROUP"])]
    lama = pd.concat([lama, pd.DataFrame([
        {"Group": "SINARMAS GROUP", "Urutan": 1, "Entitas": "PT Sinar Mas Agro Resources and Technology Tbk",
         "Kode PEFINDO": "SMAR"},
        {"Group": "BARITO PACIFIC GROUP", "Urutan": 1, "Entitas": "PT Barito Pacific Tbk", "Kode PEFINDO": "BRPT"},
        {"Group": "BARITO PACIFIC GROUP", "Urutan": 2, "Entitas": "PT Chandra Asri Pacific Tbk", "Kode PEFINDO": "TPIA"},
        {"Group": "KUSTOM GROUP", "Urutan": 1, "Entitas": "PT Kustom", "Kode PEFINDO": "KSTM"}])], ignore_index=True)
    baru = rapikan_entitas(sesuaikan_versi_lama(lama))
    assert list(baru[baru["Group"] == "SINARMAS GROUP"]["Kode PEFINDO"]) == ["ASMS"]
    assert list(baru[baru["Group"] == "BARITO PACIFIC GROUP"]["Kode PEFINDO"]) == ["BRPT"]
    assert list(baru[baru["Group"] == "BANK BTPN GROUP"]["Lembaga"]) == ["Fitch"]
    assert "KUSTOM GROUP" in set(baru["Group"])                     # isian pengguna lain tetap

"""Uji parser judul rilis & halaman PEFINDO memakai contoh teks asli dari pefindo.com (Nov 2026).
Jalankan:  python -m pytest -q   (atau python tests/test_parser.py)"""
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from core.pefindo import parse_daftar_emiten, parse_halaman_emiten, riwayat_dari_emiten  # noqa: E402
from core.teks_rating import baca_judul, cari_tanggal, parse_teks_bebas  # noqa: E402

KASUS = [
    # (judul, tingkat, rating, aksi)
    ('PEFINDO upgraded the ratings for PT Bank Rakyat Indonesia (Persero) Tbk (BBRI or the Bank) to "idA+" from"idA".',
     "perusahaan", "idA+", "Naik"),
    ('PEFINDO affirmed its rating for PT Bank Rakyat Indonesia (Persero) Tbk (BBRI or the Bank) at "idAA-", while its '
     'subordinated bond at "idA+"', "perusahaan", "idAA-", "Afirmasi"),
    ('PEFINDO upgraded the rating of PT Bank Rakyat Indonesia Tbk to "idAAA" from "idAA+".', "perusahaan", "idAAA", "Naik"),
    ("PT Bank Rakyat Indonesia (Persero) Tbk has fully repaid Surat Berharga Komersial Berkelanjutan I Seri B",
     "lain", None, ""),
    ("PEFINDO has affirmed its idAAA ratings to PT Bank Rakyat Indonesia (Persero) Tbk (BBRI) and its outstanding bonds",
     "perusahaan", "idAAA", "Afirmasi"),
    ("PT Bank Rakyat Indonesia (Persero) Tbk (BBRI)'s following debt instruments will mature in the near term: Shelf "
     "Registration Green Bond I Phase II Year 2023 Series C", "lain", None, ""),
    ("PEFINDO has assigned its idAAA rating for PT Bank Rakyat Indonesia (Persero) Tbk (BBRI)'s proposed Shelf "
     "Registered Social Bond I", "instrumen", "idAAA", "Baru"),
    ("PEFINDO has assigned its idA1+ ratings to PT Bank Rakyat Indonesia (Persero) Tbk (BBRI)'s Surat Berharga "
     "Komersial (SBK) Berkelanjutan I", "instrumen", None, "Baru"),
    ('PT Waskita Karya (Persero) Tbk rated "idBBB" with negative outlook', "perusahaan", "idBBB", "Baru"),
    ('PEFINDO assigns PT Waskita Karya (Persero) Tbk\'s proposed bond and sukuk at "idAAA(gg)"', "instrumen",
     "idAAA(gg)", "Baru"),
    ("PEFINDO has affirmed its idBBB rating for PT Waskita Karya (Persero) Tbk (WSKT)'s maturing shelf-registered (SR) "
     "Bond III", "instrumen", "idBBB", "Afirmasi"),
    ("PEFINDO has withdrawn the rating of Obligasi Berkelanjutan III Waskita Karya Tahap I Tahun 2017 Seri B",
     "instrumen", None, "Ditarik"),
    ("PEFINDO has lowered the ratings for PT Waskita Karya (Persero) Tbk (WSKT) and its shelf-registered (SR) bonds III",
     "perusahaan", None, "Turun"),
    ("PEFINDO has lowered the rating for PT Waskita Karya (Persero) Tbk (WSKT)'s shelf-registered (SR) Bond IV to idD "
     "from idCCC", "instrumen", "idD", "Turun"),
    ("PEFINDO has affirmed the rating for PT Waskita Karya (Persero) Tbk (WSKT) at idSD", "perusahaan", "idSD", "Afirmasi"),
    ("PT Waskita Karya (Persero) Tbk (WSKT)'s Bond III Year 2021 Series A (rated idAAA(gg)) of IDR722 billion",
     "instrumen", "idAAA(gg)", "Baru"),
    ("PEFINDO has raised the corporate rating of PT Waskita Karya (Persero) Tbk (WSKT) to idB from idSD",
     "perusahaan", "idB", "Naik"),
    ("PEFINDO has lowered the corporate rating of PT Waskita Karya (Persero) Tbk (WSKT) to idCCC from idB",
     "perusahaan", "idCCC", "Turun"),
    ("PEFINDO has raised the corporate rating of PT Waskita Karya (Persero) Tbk (WSKT) to idB with negative outlook",
     "perusahaan", "idB", "Naik"),
    ("PEFINDO has affirmed its idB ratings for PT Waskita Karya (Persero) Tbk (WSKT) and its Shelf Registration (SR) "
     "Bond III", "perusahaan", "idB", "Afirmasi"),
    ('Bank Mega is rated "idA+" with outlook stable', "perusahaan", "idA+", "Baru"),
    ('PEFINDO has assigned its "idAA-" corporate credit rating to PT Bank Mega Tbk (Bank Mega)', "perusahaan", "idAA-",
     "Baru"),
    ("PEFINDO has assigned its idAA- rating with stable outlook to PT Bank Mega Tbk", "perusahaan", "idAA-", "Baru"),
    # Fitch
    ("Fitch Affirms Bank Rakyat Indonesia at 'BBB'/'AAA(idn)'; Outlook Stable", "perusahaan", "AAA(idn)", "Afirmasi"),
    ("Fitch Upgrades Bank Rakyat Indonesia to 'BBB' and 'AAA(idn)'; Outlook Stable", "perusahaan", "AAA(idn)", "Naik"),
    ("Fitch Affirms Mandiri Tunas Finance at 'AA+(idn)'; Outlook Stable", "perusahaan", "AA+(idn)", "Afirmasi"),
    ("Fitch Rates Danantara Investment Management's Subordinated Debt AAA(idn)", "instrumen", "AAA(idn)", "Baru"),
    ("Fitch Publishes Surya Artha Nusantara Finance's Phase II Bond Rating of 'AA(idn)'", "instrumen", "AA(idn)", "Baru"),
    ("Fitch Assigns Indonesia Investment Authority First-Time 'BBB' and 'AAA(idn)' Ratings; Outlook Stable",
     "perusahaan", "AAA(idn)", "Baru"),
    ("Fitch Downgrades PT ABC to 'A-(idn)' from 'A(idn)'; Outlook Negative", "perusahaan", "A-(idn)", "Turun"),
]


def test_judul():
    gagal = []
    for judul, lvl, rating, aksi in KASUS:
        r = baca_judul(judul)
        if (r["tingkat"], r["rating"], r["aksi"]) != (lvl, rating, aksi):
            gagal.append((judul[:70], (r["tingkat"], r["rating"], r["aksi"]), (lvl, rating, aksi)))
    assert not gagal, "\n".join(map(str, gagal))


def test_outlook():
    assert baca_judul("Fitch Affirms X at 'AA(idn)'; Outlook Stable")["outlook"] == "Stable"
    assert baca_judul('PT Waskita Karya rated "idBBB" with negative outlook')["outlook"] == "Negative"
    assert baca_judul("PEFINDO places X on CreditWatch with negative implications")["outlook"] == "CreditWatch Negative"


def test_tanggal():
    assert cari_tanggal("3 Sep 2026")[0] == date(2026, 9, 3)
    assert cari_tanggal("rilis 15 Juli 2026 ...")[0] == date(2026, 7, 15)
    assert cari_tanggal("09/10/2026")[0] == date(2026, 10, 9)
    assert cari_tanggal("2026-10-09")[0] == date(2026, 10, 9)
    assert cari_tanggal("July 15, 2026")[0] == date(2026, 7, 15)


HTML_EMITEN = """
<html><body>
<div class="company"><h2>Waskita Karya (Persero) Tbk.</h2><span>(WSKT)</span></div>
<div class="rating-box"><span>Corporate Rating/Outlook</span> <strong>idB</strong> <span>Negative</span>
<small>Updated on 5 Oct 2026</small></div>
<ul class="tabs"><li><a>Press Release</a></li><li><a>Rating Rationale</a></li></ul>
<h4>Press Release</h4>
<table><thead><tr><th>Release Date</th><th>Title</th></tr></thead><tbody>
<tr><td>14 Mar 2022</td><td>PT Waskita Karya (Persero) Tbk rated "idBBB" with negative outlook</td></tr>
<tr><td>25 Apr 2022</td><td>PEFINDO assigns PT Waskita Karya (Persero) Tbk's proposed bond and sukuk at "idAAA(gg)"</td></tr>
<tr><td>28 Oct 2022</td><td>PEFINDO has affirmed the ratings for PT Waskita Karya (Persero) Tbk (WSKT) and its shelf-registered (SR) bonds III at idBBB</td></tr>
<tr><td>11 Jan 2023</td><td>PEFINDO has lowered the ratings for PT Waskita Karya (Persero) Tbk (WSKT) and its shelf-registered (SR) bonds III to idBB from idBBB</td></tr>
<tr><td>9 May 2023</td><td>PEFINDO has lowered the rating for PT Waskita Karya (Persero) Tbk (WSKT)'s shelf-registered (SR) Bond IV to idD from idCCC</td></tr>
<tr><td>11 Oct 2024</td><td>PEFINDO has affirmed the rating for PT Waskita Karya (Persero) Tbk (WSKT) at idSD</td></tr>
<tr><td>9 Oct 2025</td><td>PEFINDO has affirmed the rating for PT Waskita Karya (Persero) Tbk (WSKT) at idSD</td></tr>
<tr><td>10 Sep 2026</td><td>PEFINDO has raised the corporate rating of PT Waskita Karya (Persero) Tbk (WSKT) to idB from idSD</td></tr>
<tr><td>25 Sep 2026</td><td>PEFINDO has lowered the corporate rating of PT Waskita Karya (Persero) Tbk (WSKT) to idCCC from idB</td></tr>
<tr><td>28 Sep 2026</td><td>PT Waskita Karya (Persero) Tbk has fully repaid its Obligasi III Waskita Karya Tahun 2021 Seri A</td></tr>
<tr><td>29 Sep 2026</td><td>PEFINDO has raised the corporate rating of PT Waskita Karya (Persero) Tbk (WSKT) to idB with negative outlook</td></tr>
<tr><td>5 Oct 2026</td><td>PEFINDO has affirmed its idB ratings for PT Waskita Karya (Persero) Tbk (WSKT) and its Shelf Registration (SR) Bond III</td></tr>
</tbody></table>
<h4>Rating Rationale</h4>
<table><thead><tr><th>Release Date</th><th>Title</th></tr></thead><tbody>
<tr><td>16 Oct 2026</td><td>PEFINDO has affirmed its idAAA ratings to someone else</td></tr>
</tbody></table>
</body></html>
"""


def test_halaman_emiten():
    d = parse_halaman_emiten(HTML_EMITEN, "WSKT", "https://www.pefindo.com/rating-action-reports/rating-report/WSKT")
    assert d.rating == "idB" and d.outlook == "Negative" and d.tgl_update == date(2026, 10, 5), (d.rating, d.outlook, d.tgl_update)
    assert len(d.rilis) == 12, len(d.rilis)          # hanya tabel Press Release
    rows, notes = riwayat_dari_emiten(d)
    got = [(r["Tanggal Efektif"].date(), r["Rating"]) for r in rows]
    exp = [(date(2022, 3, 14), "idBBB"), (date(2022, 10, 28), "idBBB"), (date(2023, 1, 11), "idBB"),
           (date(2024, 10, 11), "idSD"), (date(2025, 10, 9), "idSD"), (date(2026, 9, 10), "idB"),
           (date(2026, 9, 25), "idCCC"), (date(2026, 9, 29), "idB"), (date(2026, 10, 5), "idB")]
    assert got == exp, got


def test_fallback_teks():
    html = HTML_EMITEN.replace("<table>", "<div>").replace("</table>", "</div>")
    for t in ("thead", "tbody", "tr", "th"):
        html = html.replace(f"<{t}>", "").replace(f"</{t}>", "")
    html = html.replace("<td>", "<p>").replace("</td>", "</p>")
    d = parse_halaman_emiten(html, "WSKT")
    assert len(d.rilis) == 12, len(d.rilis)


def test_daftar():
    html = """<table><thead><tr><th>No</th><th>Company Name</th><th>Ticker</th><th>Rating Type</th>
    <th>Industry Group</th></tr></thead><tbody>
    <tr><td>117</td><td><a href="https://www.pefindo.com/rating-action-reports/rating-report/BBRI">Bank Rakyat Indonesia (Persero) Tbk.</a></td><td>BBRI</td><td>Financial Institutions</td><td>Banking</td></tr>
    <tr><td>118</td><td><a href="/rating-action-reports/rating-report/AGRO">Bank Raya Indonesia, Tbk.</a></td><td></td><td>Financial Institutions</td><td>Banking</td></tr>
    </tbody></table>"""
    df = parse_daftar_emiten(html)
    assert list(df["Kode"]) == ["BBRI", "AGRO"], df


def test_teks_bebas():
    import pandas as pd
    cp = pd.DataFrame([
        {"Nama": "OCBC Group", "Nama Entitas": "PT Bank OCBC NISP Tbk", "Alias": "OCBC NISP; NISP", "Kode PEFINDO": "NISP"},
        {"Nama": "BRI GROUP+)", "Nama Entitas": "PT Bank Rakyat Indonesia (Persero) Tbk", "Alias": "Bank Rakyat Indonesia",
         "Kode PEFINDO": "BBRI"},
        {"Nama": "SULUTGO GROUP", "Nama Entitas": "PT Bank SulutGo", "Alias": "SulutGo; Bank Sulut", "Kode PEFINDO": "BSLT"},
    ])
    teks = """
    12 Mar 2026  Fitch Affirms OCBC NISP at 'AAA(idn)'; Outlook Stable
    2026-05-20 | Fitch Upgrades Bank SulutGo to 'A+(idn)'; Outlook Stable
    15 Jul 2026
    Fitch Affirms Bank Rakyat Indonesia at 'BBB'/'AAA(idn)'; Outlook Stable
    Fitch Rates Bank Rakyat Indonesia's Subordinated Bonds 'AA(idn)'
    """
    df = parse_teks_bebas(teks, cp, tanggal_default=date(2026, 11, 1))
    got = [(r.Tanggal.date(), r.Nama, r.Rating, r.Simpan) for r in df.itertuples()]
    assert got == [(date(2026, 3, 12), "OCBC Group", "AAA(idn)", True),
                   (date(2026, 5, 20), "SULUTGO GROUP", "A+(idn)", True),
                   (date(2026, 7, 15), "BRI GROUP+)", "AAA(idn)", True),
                   (date(2026, 11, 1), "BRI GROUP+)", "AA(idn)", False)], got


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)

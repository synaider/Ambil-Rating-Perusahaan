"""
Pengambil data rating dari situs publik PEFINDO (www.pefindo.com).

Sumber yang dipakai (tanpa login):
  * Halaman emiten   : https://www.pefindo.com/rating-action-reports/rating-report/{KODE}
      - blok "Corporate Rating/Outlook idAAA Stable Updated on 3 Sep 2026"
      - tabel "Press Release" (Release Date | Title) = riwayat aksi rating
  * Daftar emiten    : https://www.pefindo.com/rating-action-reports/rating-history
      - tabel No | Company Name | Ticker | Rating Type | Industry Group

Riwayat rating perusahaan dibangun dari judul press release (upgrade/downgrade/afirmasi/
penetapan), lalu rating terkini dari blok "Corporate Rating/Outlook" dipakai sebagai acuan
terakhir. Judul yang hanya menyangkut instrumen (obligasi, sukuk, MTN, SBK, dll.) diabaikan.

Catatan: struktur situs bisa berubah sewaktu-waktu. Parser dibuat toleran (membaca tabel HTML
maupun teks halaman), dan setiap hasil dicatat sumber URL-nya agar bisa dicek ulang.
"""
from __future__ import annotations

import re
import time
from dataclasses import dataclass, field
from datetime import date, datetime

import pandas as pd

from .engine import L_PEFINDO, norm_key, normalize_master, parse_rating
from .importer import gabung_rating
from .teks_rating import baca_judul, cari_tanggal

BASE = "https://www.pefindo.com"
URL_EMITEN = BASE + "/rating-action-reports/rating-report/{kode}"
URL_DAFTAR = BASE + "/rating-action-reports/rating-history"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/130.0 Safari/537.36 LimitCounterpartyApp/1.0",
    "Accept": "text/html,application/xhtml+xml",
    "Accept-Language": "en-US,en;q=0.9,id;q=0.8",
}

RE_HEADER = re.compile(
    r"Corporate\s+Rating\s*/\s*Outlook\s+(?P<rating>id[A-Z]{1,3}[+-]?(?:\([a-z]{2}\))?)\s*(?P<outlook>.{0,40}?)\s*"
    r"Updated\s+on\s+(?P<tgl>\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4})", re.I | re.S)
RE_TGL_BARIS = re.compile(r"^\s*(\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4})\s*$")


@dataclass
class DataEmiten:
    kode: str
    url: str
    nama: str | None = None
    rating: str | None = None
    outlook: str | None = None
    tgl_update: date | None = None
    rilis: list = field(default_factory=list)       # [(tanggal, judul)]
    peringatan: list = field(default_factory=list)


# ----------------------------------------------------------------------------
# HTTP
# ----------------------------------------------------------------------------
def ambil_html(url: str, timeout: int = 30, ulang: int = 2, sesi=None) -> str:
    import requests

    s = sesi or requests.Session()
    err = None
    for i in range(ulang + 1):
        try:
            r = s.get(url, headers=HEADERS, timeout=timeout)
            if r.status_code == 404:
                raise LookupError(f"Halaman tidak ditemukan (404): {url}")
            r.raise_for_status()
            r.encoding = r.encoding or "utf-8"
            return r.text
        except LookupError:
            raise
        except Exception as e:  # noqa: BLE001
            err = e
            if i < ulang:
                time.sleep(1.5 * (i + 1))
    raise ConnectionError(f"Gagal mengakses {url}: {err}")


# ----------------------------------------------------------------------------
# Parser halaman emiten
# ----------------------------------------------------------------------------
def _soup(html: str):
    from bs4 import BeautifulSoup
    try:
        return BeautifulSoup(html, "lxml")
    except Exception:  # noqa: BLE001
        return BeautifulSoup(html, "html.parser")


def _teks(el) -> str:
    return re.sub(r"\s+", " ", el.get_text(" ", strip=True)).strip()


def _judul_bagian(tabel) -> str:
    """Teks heading terdekat sebelum tabel (untuk membedakan Press Release vs Rating Rationale)."""
    for el in tabel.find_all_previous(limit=40):
        if el.name in ("h1", "h2", "h3", "h4", "h5", "h6", "a", "button", "li", "span", "div", "p", "strong", "b"):
            t = _teks(el)
            if 3 <= len(t) <= 40 and re.search(r"press release|rating rationale|siaran pers", t, re.I):
                return t
    return ""


def parse_halaman_emiten(html: str, kode: str = "", url: str = "") -> DataEmiten:
    soup = _soup(html)
    for t in soup(["script", "style", "noscript"]):
        t.decompose()
    d = DataEmiten(kode=kode.upper(), url=url)
    full = re.sub(r"\s+", " ", soup.get_text(" ", strip=True))

    # --- nama perusahaan: teks tepat sebelum "(KODE)"
    if kode:
        m = re.search(r"([A-Z][\w.,&'()\- ]{3,120}?)\s*\(\s*" + re.escape(kode.upper()) + r"\s*\)", full)
        if m:
            d.nama = m.group(1).strip()

    # --- rating terkini
    m = RE_HEADER.search(full)
    if m:
        d.rating = m.group("rating")
        ol = m.group("outlook").strip(" :-|")
        d.outlook = ol or None
        d.tgl_update, _ = cari_tanggal(m.group("tgl"))
    else:
        d.peringatan.append("Blok 'Corporate Rating/Outlook' tidak ditemukan.")

    # --- tabel press release
    rilis = []
    kandidat = []
    for tb in soup.find_all("table"):
        head = [_teks(c).lower() for c in tb.find_all(["th", "td"], limit=6)]
        if any("release date" in h or "tanggal" in h for h in head) and any("title" in h or "judul" in h for h in head):
            kandidat.append(tb)
    if kandidat:
        pilih = [tb for tb in kandidat if re.search(r"press release|siaran pers", _judul_bagian(tb), re.I)]
        tb = (pilih or kandidat)[0]
        for tr in tb.find_all("tr"):
            sel = tr.find_all(["td", "th"])
            if len(sel) < 2:
                continue
            tgl, _ = cari_tanggal(_teks(sel[0]))
            if tgl:
                rilis.append((tgl, _teks(sel[1])))
    if not rilis:
        # fallback: baca teks per baris  "3 Sep 2026" -> baris berikutnya = judul
        lines = [ln.strip() for ln in soup.get_text("\n").split("\n") if ln.strip()]
        mulai = next((i for i, ln in enumerate(lines) if re.fullmatch(r"press release", ln, re.I)), 0)
        tgl0 = next((i for i, ln in enumerate(lines) if i > mulai and RE_TGL_BARIS.match(ln)), mulai)
        stop = next((i for i, ln in enumerate(lines) if i > tgl0 and re.fullmatch(r"rating rationale", ln, re.I)),
                    len(lines))
        mulai = tgl0
        i = mulai
        while i < stop - 1:
            mt = RE_TGL_BARIS.match(lines[i])
            if mt:
                tgl, _ = cari_tanggal(mt.group(1))
                if tgl and not RE_TGL_BARIS.match(lines[i + 1]):
                    rilis.append((tgl, lines[i + 1]))
                    i += 2
                    continue
            i += 1
    if not rilis:
        d.peringatan.append("Tabel Press Release tidak ditemukan / kosong.")
    d.rilis = sorted(rilis, key=lambda x: x[0])
    return d


def riwayat_dari_emiten(d: DataEmiten) -> tuple[list[dict], list[str]]:
    """Ubah daftar rilis menjadi riwayat rating perusahaan (sebelum digabung per rentang)."""
    rows, notes = [], []
    terakhir = None
    for tgl, judul in d.rilis:
        info = baca_judul(judul, L_PEFINDO)
        if info["tingkat"] != "perusahaan":
            continue
        rating = info["rating"]
        if rating and not parse_rating(rating) and rating != "NR":
            rating = None
        ket = info["aksi"] or "Rilis"
        if rating is None:
            if info["aksi"] == "Afirmasi" and terakhir:
                rating = terakhir
                ket += " (rating mengikuti rilis sebelumnya)"
            else:
                if info["aksi"] in ("Naik", "Turun", "Watch"):
                    notes.append(f"{tgl:%d/%m/%Y}: aksi '{info['aksi']}' tetapi nilai rating tidak terbaca — cek manual.")
                continue
        rows.append({"Tanggal Efektif": pd.Timestamp(tgl), "Rating": rating, "Outlook": info["outlook"],
                     "Keterangan": f"{ket}: {judul[:180]}"})
        terakhir = rating
    # rating terkini dari blok header = acuan paling akhir
    if d.rating and d.tgl_update:
        t = pd.Timestamp(d.tgl_update)
        same_day = [r for r in rows if r["Tanggal Efektif"] == t]
        if same_day:
            for r in same_day:
                if norm_key(r["Rating"]) != norm_key(d.rating):
                    notes.append(f"{d.tgl_update:%d/%m/%Y}: judul rilis terbaca {r['Rating']}, blok terkini "
                                 f"{d.rating} — dipakai blok terkini.")
                r["Rating"], r["Outlook"] = d.rating, d.outlook or r["Outlook"]
        else:
            rows.append({"Tanggal Efektif": t, "Rating": d.rating, "Outlook": d.outlook,
                         "Keterangan": "Corporate Rating/Outlook terkini di halaman PEFINDO"})
        # bila ada rilis lebih baru dari tgl update dengan rating berbeda → beri catatan
        for r in rows:
            if r["Tanggal Efektif"] > t and norm_key(r["Rating"]) != norm_key(d.rating):
                notes.append(f"{r['Tanggal Efektif']:%d/%m/%Y}: rilis lebih baru ({r['Rating']}) dari blok terkini "
                             f"({d.rating}, {d.tgl_update:%d/%m/%Y}).")
    elif not rows:
        notes.append("Tidak ada rating perusahaan yang terbaca.")
    rows.sort(key=lambda r: r["Tanggal Efektif"])
    return rows, notes


def ambil_emiten(kode: str, sesi=None) -> DataEmiten:
    kode = (kode or "").strip().upper()
    url = URL_EMITEN.format(kode=kode)
    html = ambil_html(url, sesi=sesi)
    return parse_halaman_emiten(html, kode, url)


# ----------------------------------------------------------------------------
# Daftar emiten (untuk mencari kode)
# ----------------------------------------------------------------------------
def parse_daftar_emiten(html: str) -> pd.DataFrame:
    soup = _soup(html)
    rows = []
    for tb in soup.find_all("table"):
        head = [_teks(c).lower() for c in tb.find_all("th")]
        if not any("company" in h for h in head):
            continue
        idx = {h: i for i, h in enumerate(head)}
        i_nama = next((i for h, i in idx.items() if "company" in h), 1)
        i_kode = next((i for h, i in idx.items() if "ticker" in h or "kode" in h), 2)
        i_tipe = next((i for h, i in idx.items() if "rating type" in h), None)
        i_ind = next((i for h, i in idx.items() if "industry" in h), None)
        for tr in tb.find_all("tr"):
            sel = tr.find_all("td")
            if len(sel) <= max(i_nama, i_kode):
                continue
            kode = _teks(sel[i_kode])
            link = sel[i_nama].find("a") or sel[i_kode].find("a")
            if not kode and link and link.get("href"):
                kode = link["href"].rstrip("/").split("/")[-1]
            rows.append({"Nama Perusahaan": _teks(sel[i_nama]), "Kode": kode.upper(),
                         "Jenis Rating": _teks(sel[i_tipe]) if i_tipe is not None and i_tipe < len(sel) else None,
                         "Industri": _teks(sel[i_ind]) if i_ind is not None and i_ind < len(sel) else None})
    return pd.DataFrame(rows, columns=["Nama Perusahaan", "Kode", "Jenis Rating", "Industri"]).drop_duplicates()


def ambil_daftar_emiten(sesi=None) -> pd.DataFrame:
    return parse_daftar_emiten(ambil_html(URL_DAFTAR, sesi=sesi))


def cari_kode(daftar: pd.DataFrame, kata: str, n: int = 15) -> pd.DataFrame:
    import difflib
    if daftar.empty or not kata:
        return daftar.head(0)
    q = kata.lower().strip()
    d = daftar.copy()
    d["Skor"] = d.apply(lambda r: max(
        1.0 if q in str(r["Nama Perusahaan"]).lower() or q == str(r["Kode"]).lower() else 0.0,
        difflib.SequenceMatcher(None, q, str(r["Nama Perusahaan"]).lower()).ratio()), axis=1)
    return d.sort_values("Skor", ascending=False).head(n).reset_index(drop=True)


# ----------------------------------------------------------------------------
# Sinkronisasi ke master data
# ----------------------------------------------------------------------------
def terapkan_ke_master(m: dict, nama_cp: str, d: DataEmiten, rows: list[dict]) -> dict:
    """Ganti riwayat PEFINDO hasil sinkron sebelumnya (Sumber = URL pefindo) untuk counterparty ini."""
    m = {k: v.copy() for k, v in m.items()}
    rt = m["Rating"]
    k = norm_key(nama_cp)
    lama = rt["Nama"].map(norm_key).eq(k) & rt["Lembaga"].eq(L_PEFINDO) & \
        rt["Sumber"].fillna("").str.contains("pefindo.com", case=False)
    rt = rt[~lama]
    if rows:
        stamp = datetime.now().strftime("%d/%m/%Y %H:%M")
        upd = pd.Timestamp(d.tgl_update) if d.tgl_update else None
        baru = pd.DataFrame([{**r, "Nama": nama_cp, "Lembaga": L_PEFINDO, "Sumber": d.url,
                              # tiap rilis mengonfirmasi rating pada tanggalnya; rilis terakhir s.d. "Updated on"
                              "Terakhir Dicek": max(r["Tanggal Efektif"], upd) if (upd is not None and i == len(rows) - 1)
                              else r["Tanggal Efektif"],
                              "Keterangan": f"{r['Keterangan']} [sinkron {stamp}]"} for i, r in enumerate(rows)])
        rt = pd.concat([rt, baru], ignore_index=True)
    m["Rating"] = rt
    m = normalize_master(m)
    # gabungkan rentang hanya untuk baris dari pefindo agar baris manual/impor tetap utuh
    rt = m["Rating"]
    is_pef = rt["Sumber"].fillna("").str.contains("pefindo.com", case=False) & rt["Nama"].map(norm_key).eq(k)
    m["Rating"] = pd.concat([rt[~is_pef], gabung_rating(rt[is_pef])], ignore_index=True)
    return normalize_master(m)


def sinkron(m: dict, nama_list: list[str] | None = None, progress=None, jeda: float = 0.8) -> tuple[dict, pd.DataFrame]:
    """Sinkronkan rating PEFINDO untuk counterparty yang punya Kode PEFINDO.
    progress(i, n, nama) dipanggil tiap langkah (opsional)."""
    import requests

    cp = m["Counterparty"]
    cp = cp[cp["Kode PEFINDO"].notna() & (cp["Kode PEFINDO"].astype(str).str.strip() != "")]
    if nama_list is not None:
        want = {norm_key(n) for n in nama_list}
        cp = cp[cp["Nama"].map(norm_key).isin(want)]
    laporan = []
    sesi = requests.Session()
    n = len(cp)
    gagal_koneksi, sukses = 0, 0
    for i, (_, c) in enumerate(cp.iterrows(), start=1):
        kode = str(c["Kode PEFINDO"]).strip().upper()
        if gagal_koneksi >= 2 and sukses == 0:
            # dua permintaan pertama gagal tersambung → hentikan agar tidak menunggu lama
            laporan.append({"Counterparty": c["Nama"], "Kode": kode, "Status": "Dilewati",
                            "Catatan": "Dihentikan: pefindo.com tidak dapat dihubungi dari komputer ini.",
                            "URL": URL_EMITEN.format(kode=kode)})
            continue
        if progress:
            progress(i, n, c["Nama"])
        try:
            d = ambil_emiten(kode, sesi=sesi)
            rows, notes = riwayat_dari_emiten(d)
            m = terapkan_ke_master(m, c["Nama"], d, rows)
            laporan.append({"Counterparty": c["Nama"], "Kode": kode, "Nama di PEFINDO": d.nama,
                            "Rating Terkini": d.rating, "Outlook": d.outlook,
                            "Update PEFINDO": d.tgl_update, "Jumlah Aksi Rating": len(rows),
                            "Status": "OK" if rows else "Kosong",
                            "Catatan": "; ".join(d.peringatan + notes) or None, "URL": d.url})
            sukses += 1
        except Exception as e:  # noqa: BLE001
            if isinstance(e, ConnectionError):
                gagal_koneksi += 1
            laporan.append({"Counterparty": c["Nama"], "Kode": kode, "Status": "Gagal", "Catatan": str(e),
                            "URL": URL_EMITEN.format(kode=kode)})
        if i < n and (sukses or gagal_koneksi < 2):
            time.sleep(jeda)
    return m, pd.DataFrame(laporan)

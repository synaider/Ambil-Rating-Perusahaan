"""
Cek rating per tanggal: PEFINDO → Fitch → perusahaan anggota group.

Aturan untuk tiap counterparty (group) pada tanggal T:
  untuk tiap entitas group sesuai Urutan (1 = entitas utama, 2.. = anggota group):
      1. PEFINDO : rating perusahaan terakhir dengan tanggal rilis <= T
      2. Fitch   : rating terakhir di tabel Rating_Fitch dengan tanggal <= T
  rating pertama yang ditemukan dipakai.
Rating yang sudah ditarik (NR) atau yang aksi terakhirnya lebih dari ±13 bulan sebelum T
(rating tidak lagi diperbarui lembaga) dianggap tidak ada, lalu dicari ke sumber berikutnya.
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import pandas as pd

from .defaults import ENTITAS_GROUP, SUMBER_DITETAPKAN, key_group
from .engine import L_FITCH, L_PEFINDO, norm_key, parse_rating
from .pefindo import URL_EMITEN, DataEmiten, ambil_emiten, riwayat_dari_emiten
from .teks_rating import parse_tanggal

URL_FITCH_CARI = "https://www.fitchratings.com/search/?query={q}"

MASA_BERLAKU_HARI = 400   # rating tahunan + toleransi; lebih lama dari ini = tidak berlaku lagi

KOLOM_ENTITAS = ["Group", "Urutan", "Entitas", "Kode PEFINDO", "Lembaga", "Alias"]
LEMBAGA_PILIHAN = ["Otomatis", "PEFINDO", "Fitch"]
KOLOM_FITCH = ["Entitas", "Tanggal", "Rating", "Outlook", "Keterangan"]


# ----------------------------------------------------------------------------
# Pengaturan (data/pengaturan.xlsx)
# ----------------------------------------------------------------------------
def entitas_default() -> pd.DataFrame:
    rows = []
    for g, ents in ENTITAS_GROUP:
        for i, t in enumerate(ents, start=1):
            e, kode, alias = t[:3]
            lem = t[3] if len(t) > 3 else "Otomatis"
            rows.append({"Group": g, "Urutan": i, "Entitas": e, "Kode PEFINDO": kode or None, "Lembaga": lem,
                         "Alias": alias or None})
    return pd.DataFrame(rows, columns=KOLOM_ENTITAS)


def _norm_lembaga_pilihan(v) -> str:
    t = str(v or "").strip().upper()
    if "FITCH" in t:
        return "Fitch"
    if "PEFINDO" in t:
        return "PEFINDO"
    return "Otomatis"


def rapikan_entitas(df: pd.DataFrame) -> pd.DataFrame:
    d = df.copy()
    for c in KOLOM_ENTITAS:
        if c not in d.columns:
            d[c] = None
    d = d[KOLOM_ENTITAS]
    d = d[d["Group"].notna() & d["Entitas"].notna()]
    d["Group"] = d["Group"].map(key_group)
    d["Urutan"] = pd.to_numeric(d["Urutan"], errors="coerce")
    d["Kode PEFINDO"] = d["Kode PEFINDO"].map(lambda v: str(v).strip().upper() if isinstance(v, str) and v.strip()
                                              else None)
    d["Entitas"] = d["Entitas"].astype(str).str.strip()
    d["Alias"] = d["Alias"].map(lambda v: v if isinstance(v, str) and v.strip() else None)
    d["Lembaga"] = d["Lembaga"].map(_norm_lembaga_pilihan)
    d["_i"] = range(len(d))
    d = d.sort_values(["Group", "Urutan", "_i"], na_position="last").drop(columns="_i")
    d["Urutan"] = d.groupby("Group").cumcount() + 1
    return d.reset_index(drop=True)


def ke_tanggal(v):
    """Tanggal dari sel Excel / teks (2026-11-05, 05/11/2026, 5 Nov 2026, 5 November 2026)."""
    if v is None or (isinstance(v, float) and pd.isna(v)) or v is pd.NaT:
        return None
    if isinstance(v, (datetime, date, pd.Timestamp)):
        return pd.Timestamp(v).normalize()
    d = parse_tanggal(str(v))
    return pd.Timestamp(d) if d else None


def rapikan_fitch(df: pd.DataFrame) -> pd.DataFrame:
    d = df.copy()
    for c in KOLOM_FITCH:
        if c not in d.columns:
            d[c] = None
    d = d[KOLOM_FITCH]
    d = d[d["Entitas"].notna() & d["Rating"].notna()]
    d["Tanggal"] = pd.to_datetime(d["Tanggal"].map(ke_tanggal), errors="coerce").astype("datetime64[ns]")
    d["Rating"] = d["Rating"].astype(str).str.strip()
    d["Entitas"] = d["Entitas"].astype(str).str.strip()
    return d.sort_values(["Entitas", "Tanggal"], kind="stable").reset_index(drop=True)


def fitch_dari_template(rows: list[dict], tanggal: date, entitas: pd.DataFrame, nama_file: str) -> pd.DataFrame:
    """Rating '(idn)' yang tertulis di file Excel dijadikan data Fitch awal (tanggal = tanggal laporan)."""
    out = []
    for r in rows:
        p = parse_rating(r.get("Rating"))
        if not p or not p["idn"]:
            continue
        ent = entitas[entitas["Group"] == key_group(r["Nama"])]
        if ent.empty:
            continue
        out.append({"Entitas": ent.iloc[0]["Entitas"], "Tanggal": pd.Timestamp(tanggal), "Rating": r["Rating"],
                    "Outlook": None, "Keterangan": f"Dari file {nama_file}"})
    return pd.DataFrame(out, columns=KOLOM_FITCH)


def sesuaikan_versi_lama(ent: pd.DataFrame) -> pd.DataFrame:
    """Pengaturan dari versi sebelumnya (tanpa kolom Lembaga): tambahkan kolom Lembaga dan terapkan
    sumber yang sudah ditetapkan pengguna untuk group di SUMBER_DITETAPKAN."""
    e = ent.copy()
    if "Lembaga" not in e.columns:
        e["Lembaga"] = "Otomatis"
    e["_g"] = e["Group"].map(key_group)
    d = entitas_default()
    e = pd.concat([e[~e["_g"].isin(SUMBER_DITETAPKAN)].drop(columns="_g"),
                   d[d["Group"].isin(SUMBER_DITETAPKAN)]], ignore_index=True)
    return e


def muat_pengaturan(path: Path) -> tuple[pd.DataFrame, pd.DataFrame, bool]:
    """Return (entitas, fitch, disesuaikan). disesuaikan=True bila pengaturan versi lama diperbarui."""
    if path.exists():
        xls = pd.ExcelFile(path)
        ent = pd.read_excel(xls, "Entitas_Group") if "Entitas_Group" in xls.sheet_names else entitas_default()
        fit = pd.read_excel(xls, "Rating_Fitch") if "Rating_Fitch" in xls.sheet_names else pd.DataFrame(columns=KOLOM_FITCH)
        lama = "Lembaga" not in ent.columns
        if lama:
            ent = sesuaikan_versi_lama(ent)
        return rapikan_entitas(ent), rapikan_fitch(fit), lama
    return rapikan_entitas(entitas_default()), rapikan_fitch(pd.DataFrame(columns=KOLOM_FITCH)), False


def simpan_pengaturan(path: Path, ent: pd.DataFrame, fit: pd.DataFrame) -> None:
    from openpyxl.styles import Font, PatternFill
    path.parent.mkdir(parents=True, exist_ok=True)
    f2 = rapikan_fitch(fit).copy()
    f2["Tanggal"] = f2["Tanggal"].dt.date
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        rapikan_entitas(ent).to_excel(xw, sheet_name="Entitas_Group", index=False)
        f2.to_excel(xw, sheet_name="Rating_Fitch", index=False)
        for ws in xw.book.worksheets:
            for c in ws[1]:
                c.font = Font(bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="1F4E78")
            ws.freeze_panes = "A2"
            for col, w in zip("ABCDEF", (34, 12, 48, 14, 12, 50)):
                ws.column_dimensions[col].width = w
            for row in ws.iter_rows(min_row=2):
                for c in row:
                    if isinstance(c.value, (date, datetime)):
                        c.number_format = "dd/mm/yyyy"


# ----------------------------------------------------------------------------
# Cache halaman PEFINDO (data/cache_pefindo/KODE.json)
# ----------------------------------------------------------------------------
def _ke_json(d: DataEmiten) -> dict:
    return {"kode": d.kode, "url": d.url, "nama": d.nama, "rating": d.rating, "outlook": d.outlook,
            "tgl_update": d.tgl_update.isoformat() if d.tgl_update else None,
            "rilis": [[t.isoformat(), j] for t, j in d.rilis], "peringatan": d.peringatan,
            "diambil": datetime.now().isoformat(timespec="seconds")}


def _dari_json(js: dict) -> DataEmiten:
    return DataEmiten(kode=js["kode"], url=js.get("url") or URL_EMITEN.format(kode=js["kode"]), nama=js.get("nama"),
                      rating=js.get("rating"), outlook=js.get("outlook"),
                      tgl_update=date.fromisoformat(js["tgl_update"]) if js.get("tgl_update") else None,
                      rilis=[(date.fromisoformat(t), j) for t, j in js.get("rilis", [])],
                      peringatan=js.get("peringatan", []))


class SumberPefindo:
    """Mengambil halaman emiten PEFINDO dengan cache harian di disk.
    Jika pefindo.com tidak bisa diakses, memakai cache terakhir yang ada (berapa pun umurnya)."""

    def __init__(self, cache_dir: Path, paksa_ulang: bool = False, jeda: float = 0.5):
        self.dir = cache_dir
        self.dir.mkdir(parents=True, exist_ok=True)
        self.paksa = paksa_ulang
        self.jeda = jeda
        self.mem: dict[str, tuple[DataEmiten | None, str]] = {}
        self.gagal_koneksi = 0
        self.sukses = 0
        self._sesi = None
        self._terakhir = 0.0

    def _cache(self, kode: str) -> tuple[DataEmiten | None, datetime | None]:
        f = self.dir / f"{kode}.json"
        if not f.exists():
            return None, None
        try:
            js = json.loads(f.read_text(encoding="utf-8"))
            return _dari_json(js), datetime.fromisoformat(js["diambil"])
        except Exception:  # noqa: BLE001
            return None, None

    def ambil(self, kode: str) -> tuple[DataEmiten | None, str]:
        """Return (data, catatan). catatan kosong = data segar dari pefindo.com / cache hari ini."""
        import time

        kode = kode.strip().upper()
        if kode in self.mem:
            return self.mem[kode]
        cache, waktu = self._cache(kode)
        if cache is not None and not self.paksa and waktu and waktu.date() == date.today():
            self.mem[kode] = (cache, "")
            return self.mem[kode]
        offline = self.gagal_koneksi >= 2 and self.sukses == 0
        hasil: tuple[DataEmiten | None, str]
        if offline:
            hasil = (cache, f"pakai data tersimpan {waktu:%d/%m/%Y}" if cache else "pefindo.com tidak bisa diakses")
        else:
            try:
                if self._sesi is None:
                    import requests
                    self._sesi = requests.Session()
                tunggu = self.jeda - (time.time() - self._terakhir)
                if tunggu > 0:
                    time.sleep(tunggu)
                d = ambil_emiten(kode, sesi=self._sesi)
                self._terakhir = time.time()
                (self.dir / f"{kode}.json").write_text(json.dumps(_ke_json(d), ensure_ascii=False, indent=1),
                                                       encoding="utf-8")
                self.sukses += 1
                hasil = (d, "")
            except LookupError:
                hasil = (None, f"kode {kode} tidak ada di pefindo.com")
            except Exception as e:  # noqa: BLE001
                if isinstance(e, ConnectionError):
                    self.gagal_koneksi += 1
                hasil = (cache, f"pakai data tersimpan {waktu:%d/%m/%Y}" if cache else f"gagal akses pefindo.com")
        self.mem[kode] = hasil
        return hasil


# ----------------------------------------------------------------------------
# Rating pada tanggal tertentu
# ----------------------------------------------------------------------------
@dataclass
class Temuan:
    rating: str | None = None
    lembaga: str | None = None
    entitas: str | None = None
    tanggal: date | None = None
    outlook: str | None = None
    urutan: int | None = None
    url: str | None = None
    jejak: list = field(default_factory=list)     # catatan pencarian


def rating_pefindo_pada(d: DataEmiten, tgl: date) -> tuple[dict | None, str]:
    rows, _ = riwayat_dari_emiten(d)
    t = pd.Timestamp(tgl)
    valid = [r for r in rows if r["Tanggal Efektif"] <= t]
    if not valid:
        return None, "belum ada rating s.d. tanggal ini"
    r = valid[-1]
    if norm_key(r["Rating"]) in ("NR", "WD") or not parse_rating(r["Rating"]):
        return None, f"rating ditarik ({r['Tanggal Efektif']:%d/%m/%Y})"
    if (t - r["Tanggal Efektif"]).days > MASA_BERLAKU_HARI:
        return None, f"rating terakhir {r['Tanggal Efektif']:%d/%m/%Y} sudah tidak berlaku"
    return r, ""


def rating_fitch_pada(fit: pd.DataFrame, entitas: str, tgl: date) -> tuple[pd.Series | None, str]:
    if fit.empty:
        return None, "tidak ada data"
    t = pd.Timestamp(tgl)
    d = fit[(fit["Entitas"].map(norm_key) == norm_key(entitas)) & fit["Tanggal"].notna()]
    if d.empty:
        return None, "tidak ada data"
    d = d[d["Tanggal"] <= t]
    if d.empty:
        return None, "belum ada rating s.d. tanggal ini"
    r = d.iloc[-1]
    if norm_key(r["Rating"]) in ("NR", "WD") or not parse_rating(r["Rating"]):
        return None, f"rating ditarik ({r['Tanggal']:%d/%m/%Y})"
    if (t - r["Tanggal"]).days > MASA_BERLAKU_HARI:
        return None, f"rating terakhir {r['Tanggal']:%d/%m/%Y} sudah tidak berlaku"
    return r, ""


def url_fitch(nama: str) -> str:
    """Link pencarian di situs Fitch Ratings untuk nama entitas (tanpa PT/Tbk/(Persero))."""
    from urllib.parse import quote_plus
    import re
    q = re.sub(r"\b(PT|Tbk|Persero|Perseroda)\b\.?|[()]", " ", str(nama or ""), flags=re.I)
    return URL_FITCH_CARI.format(q=quote_plus(re.sub(r"\s+", " ", q).strip()))


def url_pefindo(kode: str) -> str:
    return URL_EMITEN.format(kode=str(kode).strip().upper())


def cek_group(group: str, ent: pd.DataFrame, fit: pd.DataFrame, tgl: date, pef: SumberPefindo) -> Temuan:
    rows = ent[ent["Group"] == key_group(group)].sort_values("Urutan")
    tm = Temuan()
    if rows.empty:
        tm.jejak.append("belum ada entitas untuk group ini (isi di tab Sumber Rating)")
        return tm
    for _, e in rows.iterrows():
        nama = e["Entitas"]
        kode = e["Kode PEFINDO"] if isinstance(e["Kode PEFINDO"], str) else ""
        kode = kode.strip().upper()
        pilihan = e["Lembaga"] if isinstance(e.get("Lembaga"), str) else "Otomatis"
        if pilihan == "Fitch":
            tm.jejak.append(f"{nama}: diatur pakai Fitch")
        if kode and pilihan != "Fitch":
            d, cat = pef.ambil(kode)
            if d is not None:
                r, why = rating_pefindo_pada(d, tgl)
                if r is not None:
                    tm.rating, tm.lembaga, tm.entitas = r["Rating"], L_PEFINDO, nama
                    tm.tanggal, tm.outlook, tm.urutan, tm.url = r["Tanggal Efektif"].date(), r["Outlook"], \
                        int(e["Urutan"]), d.url
                    if cat:
                        tm.jejak.append(cat)
                    return tm
                tm.jejak.append(f"PEFINDO {kode}: {why}")
            else:
                tm.jejak.append(f"PEFINDO {kode}: {cat}")
        if pilihan == "PEFINDO":
            continue
        r, why = rating_fitch_pada(fit, nama, tgl)
        if r is not None:
            tm.rating, tm.lembaga, tm.entitas = r["Rating"], L_FITCH, nama
            tm.tanggal, tm.outlook, tm.urutan = r["Tanggal"].date(), r["Outlook"] if pd.notna(r["Outlook"]) else None, \
                int(e["Urutan"])
            tm.url = url_fitch(nama)
            return tm
        tm.jejak.append(f"Fitch {nama}: {why}")
    # tidak ditemukan: link ke halaman PEFINDO entitas pertama yang punya kode, atau pencarian Fitch
    e0 = rows.iloc[0]
    if e0.get("Lembaga") == "Fitch" or not (isinstance(e0["Kode PEFINDO"], str) and e0["Kode PEFINDO"].strip()):
        tm.url = url_fitch(e0["Entitas"])
    else:
        tm.url = url_pefindo(e0["Kode PEFINDO"])
    return tm


def keterangan(tm: Temuan) -> str:
    if tm.rating is None:
        return "Tidak ditemukan → tetap pakai rating di file (" + "; ".join(tm.jejak) + ")"
    bagian = []
    if tm.urutan and tm.urutan > 1:
        bagian.append(f"dari anggota group: {tm.entitas}")
    if tm.lembaga == L_FITCH and any("diatur pakai Fitch" in j for j in tm.jejak):
        bagian.append("sumber diatur: Fitch")
    elif tm.lembaga == L_FITCH and any(j.startswith("PEFINDO") for j in tm.jejak):
        bagian.append("PEFINDO tidak ada → Fitch")
    bagian += [j for j in tm.jejak if "data tersimpan" in j]
    return "; ".join(bagian)

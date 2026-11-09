"""
Engine perhitungan Usulan Limit Counterparty Risiko Kredit per periode.

Logika perhitungan mengikuti file Excel "Limit_Countreparty":

NON MANAJER INVESTASI (rating eksternal PEFINDO / Fitch)
    (D) Total Aset Investasi
    (A) % batasan investasi umum           -> Parameter NONMI_BATAS_UMUM
    (B) % bobot per rating eksternal        -> Bobot_Rating (Jenis = Eksternal)
    (C) = (A) x (B)
    (L) = ROUND((C) x (D), 0)               -> Limit Nominal

MANAJER INVESTASI (rating internal)
    (A) Arahan maksimal investasi di RD     -> Parameter MI_ARAHAN_RD
    (B) Total Aset Investasi
    (C) = (A) x (B)
    (D) % batasan umum                      -> Parameter MI_BATAS_UMUM
    (E) % bobot per rating internal         -> Bobot_Rating (Jenis = Internal)
    (F) = (D) x (E)
    (L) = ROUND((F) x (C), 0)               -> Limit Nominal

Semua data disimpan sebagai RIWAYAT bertanggal. Untuk periode yang dipilih,
engine mengambil record TERAKHIR yang tanggalnya <= periode tersebut.
Rating disimpan per lembaga (PEFINDO / Fitch / Internal), sehingga sumber rating
bisa dipilih saat perhitungan.
"""
from __future__ import annotations

import calendar
import re
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import ROUND_HALF_UP, Decimal
from io import BytesIO
from pathlib import Path

import pandas as pd

# --------------------------------------------------------------------------
# Struktur master data (satu file Excel, satu sheet per tabel)
# --------------------------------------------------------------------------
SHEETS: dict[str, list[str]] = {
    "Counterparty": ["No", "Nama", "Kategori", "Lembaga Utama", "Kode PEFINDO", "Nama Entitas", "Alias", "Keterangan"],
    "Rating": ["Nama", "Lembaga", "Tanggal Efektif", "Rating", "Outlook", "Terakhir Dicek", "Sumber", "Keterangan"],
    "Status_Limit": ["Nama", "Tanggal Efektif", "Status", "Limit Nominal", "Keterangan"],
    "Total_Aset": ["Tanggal", "Total Aset Investasi", "Keterangan"],
    "Parameter": ["Kode", "Tanggal Efektif", "Nilai", "Keterangan"],
    "Bobot_Rating": ["Jenis", "Tanggal Efektif", "Urutan", "Rating", "Bobot"],
    "Eksposur": ["Tanggal", "Nama", "Aktual Nominal"],
    "Catatan": ["Bagian", "Teks"],
}
DATE_COLS = {
    "Rating": ["Tanggal Efektif", "Terakhir Dicek"],
    "Status_Limit": ["Tanggal Efektif"],
    "Total_Aset": ["Tanggal"],
    "Parameter": ["Tanggal Efektif"],
    "Bobot_Rating": ["Tanggal Efektif"],
    "Eksposur": ["Tanggal"],
}
NUM_COLS = {
    "Counterparty": ["No"],
    "Total_Aset": ["Total Aset Investasi"],
    "Parameter": ["Nilai"],
    "Bobot_Rating": ["Urutan", "Bobot"],
    "Status_Limit": ["Limit Nominal"],
    "Eksposur": ["Aktual Nominal"],
}

KAT_NON_MI = "Non MI"
KAT_MI = "MI"
JENIS_EKSTERNAL = "Eksternal"
JENIS_INTERNAL = "Internal"

L_PEFINDO = "PEFINDO"
L_FITCH = "Fitch"
L_INTERNAL = "Internal"
LEMBAGA_LIST = [L_PEFINDO, L_FITCH, L_INTERNAL]

STATUS_RUMUS = "Rumus"            # limit dihitung dari rumus (default)
STATUS_OVERRIDE = "Override"      # limit diisi manual
STATUS_KHUSUS = "Khusus"          # limit manual + disorot hijau (mis. BRI +)
STATUS_NONAKTIF = "Non-aktif"     # baris abu-abu, tanpa tanda "v"
STATUS_LIST = [STATUS_RUMUS, STATUS_OVERRIDE, STATUS_KHUSUS, STATUS_NONAKTIF]

P_NONMI_A = "NONMI_BATAS_UMUM"
P_MI_ARAHAN = "MI_ARAHAN_RD"
P_MI_D = "MI_BATAS_UMUM"
PARAM_LABEL = {
    P_NONMI_A: "(A) % batasan investasi terhadap Total Aset Investasi (Non MI)",
    P_MI_ARAHAN: "(A) Arahan maksimal investasi di Reksa Dana (MI)",
    P_MI_D: "(D) % batasan investasi terhadap poin (C) (MI)",
}

# Pilihan sumber rating untuk counterparty Non MI
MODE_UTAMA = "Lembaga utama (seperti Excel)"
MODE_PEFINDO = "PEFINDO saja"
MODE_FITCH = "Fitch saja"
MODE_PEF_FITCH = "PEFINDO → cadangan Fitch"
MODE_FITCH_PEF = "Fitch → cadangan PEFINDO"
MODE_TERENDAH = "Terendah PEFINDO/Fitch (konservatif)"
MODE_LIST = [MODE_UTAMA, MODE_PEF_FITCH, MODE_FITCH_PEF, MODE_PEFINDO, MODE_FITCH, MODE_TERENDAH]

# Skala lengkap untuk mendeteksi naik/turun rating (termasuk di luar tabel bobot)
SKALA = ["AAA", "AA+", "AA", "AA-", "A+", "A", "A-", "BBB+", "BBB", "BBB-",
         "BB+", "BB", "BB-", "B+", "B", "B-", "CCC", "CC", "C", "SD", "RD", "D"]
# Akhiran di belakang "idXXX" yang dianggap setara dengan rating polos (sesuai Excel: idAAA(gg)).
# Tambahkan mis. "(sy)" bila rating sukuk ingin ikut dipetakan.
AKHIRAN_DIIZINKAN = ("(gg)",)

_RATING_RE = re.compile(
    r"^(?P<pre>id)?(?P<core>AAA|AA|A|BBB|BB|B|CCC|CC|C|SD|RD|D)(?P<sign>[+-])?(?P<suf>(?:\([a-z]+\))*)$",
    re.IGNORECASE,
)


def norm_key(s) -> str:
    """Kunci pencocokan nama: trim, spasi ganda dirapatkan, huruf besar."""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return re.sub(r"\s+", " ", str(s)).strip().upper()


def kosong(v) -> bool:
    return v is None or (isinstance(v, float) and pd.isna(v)) or (isinstance(v, str) and not v.strip()) or v is pd.NaT


def excel_round(x: float, digits: int = 0) -> float:
    """ROUND ala Excel (half away from zero, presisi 15 digit)."""
    if x is None or pd.isna(x):
        return x
    d = Decimal(f"{x:.15g}")
    q = Decimal(1).scaleb(-digits)
    return float(d.quantize(q, rounding=ROUND_HALF_UP))


# --------------------------------------------------------------------------
# Rating
# --------------------------------------------------------------------------
def parse_rating(r) -> dict | None:
    if kosong(r):
        return None
    s = re.sub(r"\s+", "", str(r))
    m = _RATING_RE.match(s)
    if not m:
        return None
    sufs = [x.lower() for x in re.findall(r"\([a-zA-Z]+\)", m.group("suf") or "")]
    return {
        "core": m.group("core").upper() + (m.group("sign") or ""),
        "id": bool(m.group("pre")),
        "idn": "(idn)" in sufs,
        "extra": [x for x in sufs if x != "(idn)"],
    }


def tier_eksternal(rating, tiers: list[str]) -> str | None:
    """Petakan rating ke tier bobot. Sama persis dengan rumus Excel:
    'idAAA', 'idAAA(gg)', dan 'AAA(idn)' -> tier 'AAA'. Selain itu -> None (NA)."""
    p = parse_rating(rating)
    if not p or p["core"] not in tiers:
        return None
    if p["id"] and not p["idn"] and len(p["extra"]) <= 1 and all(x in AKHIRAN_DIIZINKAN for x in p["extra"]):
        return p["core"]
    if not p["id"] and p["idn"] and not p["extra"]:
        return p["core"]
    return None


def lembaga_dari_rating(rating) -> str:
    p = parse_rating(rating)
    if not p:
        return ""
    if p["id"]:
        return L_PEFINDO
    if p["idn"]:
        return L_FITCH
    return ""


def norm_lembaga(v) -> str:
    k = norm_key(v)
    if not k:
        return ""
    if "PEFINDO" in k or k in ("PEF",):
        return L_PEFINDO
    if "FITCH" in k:
        return L_FITCH
    if "INTERN" in k:
        return L_INTERNAL
    return str(v).strip()


def rank_rating(rating, internal_labels: list[str] | None = None) -> int | None:
    """Peringkat numerik (0 = terbaik) untuk deteksi naik/turun."""
    if internal_labels:
        k = norm_key(rating)
        for i, lab in enumerate(internal_labels):
            if norm_key(lab) == k:
                return i
    p = parse_rating(rating)
    if p and p["core"] in SKALA:
        return SKALA.index(p["core"])
    return None


# --------------------------------------------------------------------------
# Load / save master
# --------------------------------------------------------------------------
def empty_master() -> dict[str, pd.DataFrame]:
    return normalize_master({k: pd.DataFrame(columns=v) for k, v in SHEETS.items()})


def _to_str_or_none(v):
    if kosong(v):
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def normalize_master(m: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    out = {}
    for sheet, cols in SHEETS.items():
        df = m.get(sheet)
        df = pd.DataFrame(columns=cols) if df is None else df.copy()
        df.columns = [str(c).strip() for c in df.columns]
        for c in cols:
            if c not in df.columns:
                df[c] = None
        df = df[cols]
        for c in DATE_COLS.get(sheet, []):
            df[c] = pd.to_datetime(df[c], errors="coerce", dayfirst=True).dt.normalize()
            df[c] = df[c].astype("datetime64[ns]")
        for c in NUM_COLS.get(sheet, []):
            df[c] = pd.to_numeric(df[c], errors="coerce").astype("float64")
        for c in cols:
            if c not in DATE_COLS.get(sheet, []) and c not in NUM_COLS.get(sheet, []):
                df[c] = df[c].map(_to_str_or_none).astype(object)
        df = df.dropna(how="all").reset_index(drop=True)
        out[sheet] = df

    # Lengkapi kolom Lembaga yang kosong
    rt = out["Rating"]
    if not rt.empty:
        kat = dict(zip(out["Counterparty"]["Nama"].map(norm_key), out["Counterparty"]["Kategori"]))
        lem = []
        for _, r in rt.iterrows():
            v = norm_lembaga(r["Lembaga"])
            if not v:
                v = L_INTERNAL if kat.get(norm_key(r["Nama"])) == KAT_MI else lembaga_dari_rating(r["Rating"])
            lem.append(v or None)
        rt["Lembaga"] = pd.Series(lem, index=rt.index, dtype=object)
    cp = out["Counterparty"]
    if not cp.empty:
        cp["Lembaga Utama"] = cp.apply(
            lambda r: norm_lembaga(r["Lembaga Utama"]) or (L_INTERNAL if r["Kategori"] == KAT_MI else L_PEFINDO), axis=1
        ).astype(object)
        cp["Kode PEFINDO"] = cp["Kode PEFINDO"].map(lambda v: v.upper() if isinstance(v, str) else v).astype(object)
    return out


def load_master(src) -> dict[str, pd.DataFrame]:
    """src: path atau file-like (.xlsx)."""
    if isinstance(src, (bytes, bytearray)):
        src = BytesIO(src)
    xls = pd.ExcelFile(src)
    m = {}
    for sheet in SHEETS:
        if sheet in xls.sheet_names:
            m[sheet] = pd.read_excel(xls, sheet_name=sheet)
    return normalize_master(m)


def validate_master(m: dict[str, pd.DataFrame]) -> list[str]:
    errs = []
    cp = m["Counterparty"]
    if cp.empty:
        errs.append("Sheet Counterparty masih kosong.")
    bad_kat = cp[~cp["Kategori"].isin([KAT_NON_MI, KAT_MI])]
    if not bad_kat.empty:
        errs.append(f"Kategori harus '{KAT_NON_MI}' atau '{KAT_MI}': " + ", ".join(bad_kat["Nama"].astype(str).head(5)))
    dup = cp[cp["Nama"].map(norm_key).duplicated()]
    if not dup.empty:
        errs.append("Nama counterparty ganda: " + ", ".join(dup["Nama"].astype(str).head(5)))
    ov = m["Status_Limit"]
    bad_st = ov[~ov["Status"].isin(STATUS_LIST)]
    if not bad_st.empty:
        errs.append("Status di sheet Status_Limit harus salah satu dari: " + ", ".join(STATUS_LIST))
    bad_l = m["Rating"][~m["Rating"]["Lembaga"].isin(LEMBAGA_LIST)]
    if not bad_l.empty:
        errs.append("Lembaga di sheet Rating harus salah satu dari: " + ", ".join(LEMBAGA_LIST))
    for sh in ("Rating", "Status_Limit", "Parameter", "Bobot_Rating"):
        n = m[sh]["Tanggal Efektif"].isna().sum()
        if n:
            errs.append(f"Ada {n} baris di sheet {sh} tanpa Tanggal Efektif.")
    known = set(cp["Nama"].map(norm_key))
    for sh in ("Rating", "Status_Limit", "Eksposur"):
        asing = sorted(set(m[sh]["Nama"].dropna().map(norm_key)) - known)
        if asing:
            errs.append(f"Sheet {sh} memuat nama yang tidak ada di Counterparty: " + ", ".join(asing[:5]))
    return errs


PETUNJUK = {
    "Counterparty": "Daftar counterparty & urutan tampil (No). Kategori: 'Non MI' (rating eksternal) atau 'MI' "
                    "(Manajer Investasi, rating internal). Lembaga Utama: PEFINDO / Fitch / Internal (dipakai bila "
                    "sumber rating = 'Ikuti lembaga utama'). Kode PEFINDO: kode emiten di pefindo.com untuk sinkron "
                    "otomatis (mis. BBRI). Nama Entitas & Alias (pisahkan dengan ;) dipakai untuk mencocokkan teks "
                    "rilis Fitch/PEFINDO ke counterparty.",
    "Rating": "RIWAYAT rating per lembaga. Satu baris = satu rentang rating yang sama, mulai Tanggal Efektif s.d. "
              "Terakhir Dicek (tanggal terakhir rating itu dikonfirmasi/afirmasi). Untuk suatu periode, aplikasi "
              "memakai baris dengan Tanggal Efektif terakhir <= periode untuk tiap lembaga. Format: idAAA / "
              "idAAA(gg) (PEFINDO), AAA(idn) (Fitch), Sangat Bagus / Bagus / Cukup / Kurang (Internal MI). "
              "Rating 'NR' = rating ditarik.",
    "Status_Limit": "RIWAYAT cara penetapan limit per counterparty. Status: Rumus (default), Override (limit manual), "
                    "Khusus (limit manual, disorot hijau, mis. BRI), Non-aktif (abu-abu, tanpa tanda 'v'; Limit "
                    "kosong = pakai hasil rumus). Counterparty tanpa baris di sini = Rumus.",
    "Total_Aset": "Total Aset Investasi (miliar Rp) per tanggal posisi. Aplikasi memakai tanggal terakhir <= periode.",
    "Parameter": "Persentase umum (riwayat). Kode: NONMI_BATAS_UMUM (A Non MI, mis. 0.10), MI_ARAHAN_RD (A MI, "
                 "mis. 0.15), MI_BATAS_UMUM (D MI, mis. 0.10). Isi sebagai pecahan (0.10 = 10%).",
    "Bobot_Rating": "Bobot (B) per rating. Jenis 'Eksternal' (AAA s/d BBB) dan 'Internal' (Sangat Bagus s/d Kurang). "
                    "Satu set bobot = semua baris dengan Jenis & Tanggal Efektif yang sama. Urutan: 1 = terbaik.",
    "Eksposur": "Aktual Nominal (miliar Rp) per tanggal posisi. Aplikasi memakai snapshot tanggal terakhir <= periode. "
                "Boleh dikosongkan.",
    "Catatan": "Catatan kaki laporan. Bagian: 'Non MI' atau 'MI'.",
}


def master_to_excel_bytes(m: dict[str, pd.DataFrame]) -> bytes:
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as xw:
        petunjuk = pd.DataFrame({"Sheet": list(PETUNJUK.keys()), "Isi / cara mengisi": list(PETUNJUK.values())})
        petunjuk.to_excel(xw, sheet_name="Petunjuk", index=False)
        for sheet in SHEETS:
            df = m[sheet].copy()
            for c in DATE_COLS.get(sheet, []):
                df[c] = df[c].dt.date
            df.to_excel(xw, sheet_name=sheet, index=False)
        wb = xw.book
        hdr_fill = PatternFill("solid", fgColor="1F4E78")
        for ws in wb.worksheets:
            for cell in ws[1]:
                cell.font = Font(name="Arial", bold=True, color="FFFFFF")
                cell.fill = hdr_fill
                cell.alignment = Alignment(vertical="center", wrap_text=True)
            ws.freeze_panes = "A2"
            for i, col in enumerate(ws.iter_cols(min_row=1, max_row=min(ws.max_row, 300)), start=1):
                width = max((len(str(c.value)) for c in col if c.value is not None), default=8)
                ws.column_dimensions[get_column_letter(i)].width = min(max(width + 2, 10), 70)
            for row in ws.iter_rows(min_row=2):
                for c in row:
                    c.font = Font(name="Arial", size=10)
                    if isinstance(c.value, (date, datetime)):
                        c.number_format = "dd/mm/yyyy"
        wb["Petunjuk"].column_dimensions["B"].width = 120
        for row in wb["Petunjuk"].iter_rows(min_row=2):
            row[1].alignment = Alignment(wrap_text=True, vertical="top")
    return bio.getvalue()


def save_master(m: dict[str, pd.DataFrame], path: str | Path) -> None:
    Path(path).write_bytes(master_to_excel_bytes(m))


# --------------------------------------------------------------------------
# Helper as-of
# --------------------------------------------------------------------------
def _ts(d) -> pd.Timestamp:
    return pd.Timestamp(d).normalize()


def latest_per_key(df: pd.DataFrame, date_col: str, as_of, key_cols=("Nama",)) -> dict[tuple, pd.Series]:
    """Record terakhir per kunci dengan tanggal <= as_of (baris terbawah menang jika tanggal sama)."""
    if isinstance(key_cols, str):
        key_cols = (key_cols,)
    as_of = _ts(as_of)
    d = df[df[date_col].notna() & (df[date_col] <= as_of)].copy()
    if d.empty:
        return {}
    d["_k"] = d[list(key_cols)].apply(lambda r: tuple(norm_key(v) for v in r), axis=1)
    d["_i"] = range(len(d))
    d = d.sort_values([date_col, "_i"], kind="stable")
    last = d.groupby("_k", sort=False).tail(1)
    return {r["_k"]: r for _, r in last.iterrows()}


def end_of_month(y: int, m: int) -> date:
    return date(y, m, calendar.monthrange(y, m)[1])


def data_date_range(m: dict[str, pd.DataFrame]) -> tuple[date | None, date | None]:
    ds = []
    for sheet, cols in DATE_COLS.items():
        if sheet in ("Parameter", "Bobot_Rating"):
            continue
        for c in cols:
            ds += list(m[sheet][c].dropna())
    if not ds:
        return None, None
    return min(ds).date(), max(ds).date()


# --------------------------------------------------------------------------
# Perhitungan
# --------------------------------------------------------------------------
@dataclass
class Hasil:
    as_of: date
    mode: str = MODE_UTAMA
    total_aset: float | None = None
    total_aset_tgl: date | None = None
    params: dict = field(default_factory=dict)          # kode -> (nilai, tgl efektif)
    tiers_ext: list = field(default_factory=list)       # [(tier, bobot)]
    tiers_int: list = field(default_factory=list)       # [(label, bobot)]
    bobot_ext_tgl: date | None = None
    bobot_int_tgl: date | None = None
    limit_ext: dict = field(default_factory=dict)       # tier -> limit
    limit_int: dict = field(default_factory=dict)       # label -> limit
    mi_c: float | None = None                           # (C) MI
    eksposur_tgl: date | None = None
    rows: pd.DataFrame = field(default_factory=pd.DataFrame)
    tidak_tercakup: list = field(default_factory=list)  # counterparty tanpa data s.d. periode
    errors: list = field(default_factory=list)
    warnings: list = field(default_factory=list)
    catatan: dict = field(default_factory=dict)          # bagian -> [teks]

    @property
    def non_mi(self) -> pd.DataFrame:
        return self.rows[self.rows["Kategori"] == KAT_NON_MI].reset_index(drop=True) if not self.rows.empty else self.rows

    @property
    def mi(self) -> pd.DataFrame:
        return self.rows[self.rows["Kategori"] == KAT_MI].reset_index(drop=True) if not self.rows.empty else self.rows

    def total(self, kat: str | None = None, col: str = "Limit Nominal") -> float:
        df = self.rows if kat is None else self.rows[self.rows["Kategori"] == kat]
        if df.empty:
            return 0.0
        return float(pd.to_numeric(df[col], errors="coerce").fillna(0).sum())


ROW_COLS = ["No", "Kategori", "Nama", "Rating", "Lembaga", "Outlook", "Tgl Rating", "Terakhir Dicek", "Hari Sejak Dicek",
            "Rating Lembaga Lain", "Tier", "Bobot", "Limit Rumus", "Limit Nominal", "Sumber Limit", "Manual", "Aktif", "Tanda",
            "Sorot", "Aktual Nominal", "Sisa Limit", "Utilisasi", "Sumber Rating", "Catatan Cek"]


def _rating_valid(r) -> bool:
    return r is not None and parse_rating(r["Rating"]) is not None


def pilih_rating(cands: dict[str, pd.Series], mode: str, utama: str) -> tuple[pd.Series | None, str]:
    """Pilih satu record rating dari kandidat per lembaga sesuai mode. Return (record, keterangan)."""
    pef, fit = cands.get(L_PEFINDO), cands.get(L_FITCH)
    pef_ok, fit_ok = _rating_valid(pef), _rating_valid(fit)

    def first(*opts):
        for o in opts:
            if o is not None and _rating_valid(o):
                return o
        # tidak ada yang valid → kembalikan record pertama yg ada (mis. NR) agar terlihat
        for o in opts:
            if o is not None:
                return o
        return None

    if mode == MODE_PEFINDO:
        return pef, ""
    if mode == MODE_FITCH:
        return fit, ""
    if mode == MODE_PEF_FITCH:
        r = first(pef, fit)
        return r, ("PEFINDO kosong → pakai Fitch" if r is fit and fit is not None and not pef_ok else "")
    if mode == MODE_FITCH_PEF:
        r = first(fit, pef)
        return r, ("Fitch kosong → pakai PEFINDO" if r is pef and pef is not None and not fit_ok else "")
    if mode == MODE_TERENDAH:
        if pef_ok and fit_ok:
            rp, rf = rank_rating(pef["Rating"]), rank_rating(fit["Rating"])
            return (fit, "Fitch lebih rendah") if (rf or 0) > (rp or 0) else (pef, "")
        return first(pef, fit), ""
    # MODE_UTAMA
    if utama == L_FITCH:
        r = first(fit, pef)
        return r, ("Fitch kosong → pakai PEFINDO" if r is pef and pef is not None and not fit_ok else "")
    r = first(pef, fit)
    return r, ("PEFINDO kosong → pakai Fitch" if r is fit and fit is not None and not pef_ok else "")


def hitung(m: dict[str, pd.DataFrame], as_of, mode: str = MODE_UTAMA, batas_umur_rating: int = 365) -> Hasil:
    as_of_ts = _ts(as_of)
    res = Hasil(as_of=as_of_ts.date(), mode=mode)

    # --- Total aset
    ta = m["Total_Aset"]
    ta = ta[ta["Tanggal"].notna() & (ta["Tanggal"] <= as_of_ts) & ta["Total Aset Investasi"].notna()]
    if ta.empty:
        res.errors.append("Belum ada data Total Aset Investasi dengan tanggal <= periode yang dipilih. "
                          "Tambahkan di menu Master Data → Total_Aset.")
    else:
        r = ta.sort_values("Tanggal", kind="stable").iloc[-1]
        res.total_aset = float(r["Total Aset Investasi"])
        res.total_aset_tgl = r["Tanggal"].date()

    # --- Parameter (jika periode lebih awal dari data parameter pertama, pakai yang paling awal + peringatan)
    pr = m["Parameter"]
    pl = latest_per_key(pr, "Tanggal Efektif", as_of_ts, key_cols=("Kode",))
    for kode in (P_NONMI_A, P_MI_ARAHAN, P_MI_D):
        r = pl.get((norm_key(kode),))
        if r is None:
            semua = pr[(pr["Kode"].map(norm_key) == norm_key(kode)) & pr["Tanggal Efektif"].notna()]
            if not semua.empty:
                r = semua.sort_values("Tanggal Efektif", kind="stable").iloc[0]
                res.warnings.append(f"Parameter {kode} paling awal berlaku sejak {r['Tanggal Efektif']:%d/%m/%Y}; "
                                    f"nilai tersebut dipakai untuk periode ini.")
        if r is None or pd.isna(r["Nilai"]):
            res.errors.append(f"Parameter {kode} belum ada di master data.")
        else:
            res.params[kode] = (float(r["Nilai"]), r["Tanggal Efektif"].date())

    # --- Bobot rating
    bw = m["Bobot_Rating"]
    for jenis in (JENIS_EKSTERNAL, JENIS_INTERNAL):
        semua = bw[(bw["Jenis"].map(norm_key) == norm_key(jenis)) & bw["Tanggal Efektif"].notna()]
        d = semua[semua["Tanggal Efektif"] <= as_of_ts]
        if d.empty and not semua.empty:
            d = semua[semua["Tanggal Efektif"] == semua["Tanggal Efektif"].min()]
            res.warnings.append(f"Bobot rating {jenis} paling awal berlaku sejak {d['Tanggal Efektif'].min():%d/%m/%Y}; "
                                f"bobot tersebut dipakai untuk periode ini.")
        if d.empty:
            res.errors.append(f"Bobot rating {jenis} belum ada di master data.")
            continue
        tgl = d["Tanggal Efektif"].max()
        d = d[d["Tanggal Efektif"] == tgl].sort_values("Urutan", kind="stable")
        pairs = [(str(r["Rating"]).strip(), float(r["Bobot"]) if pd.notna(r["Bobot"]) else 0.0) for _, r in d.iterrows()]
        if jenis == JENIS_EKSTERNAL:
            res.tiers_ext, res.bobot_ext_tgl = [(t.upper(), b) for t, b in pairs], tgl.date()
        else:
            res.tiers_int, res.bobot_int_tgl = pairs, tgl.date()

    # --- Catatan
    for _, r in m["Catatan"].iterrows():
        if r["Teks"]:
            res.catatan.setdefault(r["Bagian"] or KAT_NON_MI, []).append(r["Teks"])

    if res.errors:
        res.rows = pd.DataFrame(columns=ROW_COLS)
        return res

    TA = res.total_aset
    a = res.params[P_NONMI_A][0]
    for tier, b in res.tiers_ext:          # urutan operasi sama dgn Excel: ROUND((A*B)*D, 0)
        res.limit_ext[tier] = excel_round((a * b) * TA, 0)
    res.mi_c = res.params[P_MI_ARAHAN][0] * TA
    dmi = res.params[P_MI_D][0]
    for lab, e in res.tiers_int:
        res.limit_int[lab] = excel_round((dmi * e) * res.mi_c, 0)

    tiers = [t for t, _ in res.tiers_ext]
    bobot_ext = dict(res.tiers_ext)
    int_lookup = {norm_key(l): (l, e) for l, e in res.tiers_int}

    rat = latest_per_key(m["Rating"], "Tanggal Efektif", as_of_ts, key_cols=("Nama", "Lembaga"))
    rat_by_name: dict[str, dict[str, pd.Series]] = {}
    for (nm, lem), r in rat.items():
        rat_by_name.setdefault(nm, {})[norm_lembaga(lem) or lem] = r
    ovr = {k[0]: v for k, v in latest_per_key(m["Status_Limit"], "Tanggal Efektif", as_of_ts).items()}

    # --- Eksposur: snapshot tanggal terakhir <= periode
    ek = m["Eksposur"]
    ek = ek[ek["Tanggal"].notna() & (ek["Tanggal"] <= as_of_ts)]
    eks = {}
    if not ek.empty:
        tgl = ek["Tanggal"].max()
        res.eksposur_tgl = tgl.date()
        for _, r in ek[ek["Tanggal"] == tgl].iterrows():
            if pd.notna(r["Aktual Nominal"]):
                eks[norm_key(r["Nama"])] = float(r["Aktual Nominal"])

    known = set()
    out = []
    cp = m["Counterparty"].copy()
    cp["_i"] = range(len(cp))
    cp["_kat"] = (cp["Kategori"] == KAT_MI).astype(int)
    cp = cp.sort_values(["_kat", "No", "_i"], kind="stable", na_position="last")
    for _, c in cp.iterrows():
        k = norm_key(c["Nama"])
        known.add(k)
        kat = c["Kategori"]
        cands = rat_by_name.get(k, {})
        o = ovr.get(k)
        o_status = o["Status"] if o is not None else None
        if not cands and (o is None or o_status == STATUS_RUMUS):
            res.tidak_tercakup.append(c["Nama"])
            continue

        notes = []
        if kat == KAT_MI:
            r = cands.get(L_INTERNAL)
            if r is None:
                r = next(iter(cands.values()), None)
            ket_pilih, lain = "", None
        else:
            r, ket_pilih = pilih_rating(cands, mode, c["Lembaga Utama"] or L_PEFINDO)
            lain_l = [x for x in (L_PEFINDO, L_FITCH) if x in cands and (r is None or cands[x] is not r)]
            lain = "; ".join(f"{x}: {cands[x]['Rating']}" for x in lain_l) or None
            if ket_pilih:
                notes.append(ket_pilih)
        rating = r["Rating"] if r is not None else None

        tier, bobot, limit_rumus = None, None, None
        if kat == KAT_MI:
            hit = int_lookup.get(norm_key(rating)) if rating else None
            if hit:
                tier, bobot = hit
                limit_rumus = res.limit_int[tier]
        else:
            tier = tier_eksternal(rating, tiers) if rating else None
            if tier:
                bobot = bobot_ext[tier]
                limit_rumus = res.limit_ext[tier]

        aktif, tanda, sorot = True, "v", ""
        sumber = STATUS_RUMUS
        limit = limit_rumus
        manual = False
        if o is not None and o_status != STATUS_RUMUS:
            v = o["Limit Nominal"]
            if o_status == STATUS_NONAKTIF:
                aktif, tanda, sorot = False, "", "abu"
                sumber = STATUS_NONAKTIF
                if pd.notna(v):
                    limit, manual = float(v), True
            elif pd.notna(v):
                limit, manual = float(v), True
                sumber = o_status
                if o_status == STATUS_KHUSUS:
                    sorot = "hijau"
            else:
                notes.append("Override tanpa nilai limit → dipakai rumus")

        if rating is None:
            notes.append("Belum ada data rating" + ("" if kat == KAT_MI or mode in (MODE_UTAMA, MODE_TERENDAH,
                                                                                   MODE_PEF_FITCH, MODE_FITCH_PEF)
                                                     else f" ({mode})"))
        elif norm_key(rating) in ("NR", "WD", "WR"):
            notes.append("Rating ditarik (NR)")
        elif tier is None:
            notes.append("Rating di luar tabel bobot (NA)")
        if limit is None:
            notes.append("Limit NA")
        umur = None
        tgl_rating = tgl_cek = None
        outlook = sumber_rating = None
        if r is not None:
            tgl_rating = r["Tanggal Efektif"].date()
            akhir = r["Terakhir Dicek"] if pd.notna(r["Terakhir Dicek"]) else r["Tanggal Efektif"]
            akhir = max(min(akhir, as_of_ts), r["Tanggal Efektif"])
            tgl_cek = akhir.date()
            umur = max((as_of_ts - akhir).days, 0)   # hari sejak terakhir dikonfirmasi
            if batas_umur_rating and umur > batas_umur_rating and kat != KAT_MI:
                notes.append(f"Rating tidak dikonfirmasi > {batas_umur_rating} hari")
            outlook = r["Outlook"]
            sumber_rating = r["Sumber"]
            if isinstance(outlook, str) and re.search(r"negati|watch", outlook, re.I):
                notes.append(f"Outlook {outlook}")
        if sumber in (STATUS_OVERRIDE, STATUS_KHUSUS) and limit_rumus is not None and limit != limit_rumus:
            notes.append(f"Limit manual (rumus: {limit_rumus:,.0f})")

        aktual = eks.get(k)
        sisa = util = None
        if aktual is not None and limit is not None:
            sisa = limit - aktual
            util = aktual / limit if limit else None
            if aktual > limit:
                notes.append("Aktual MELEBIHI limit")
        if not aktif:
            notes.append("Non-aktif")

        lembaga = (r["Lembaga"] if r is not None and r["Lembaga"] else None) or (
            L_INTERNAL if kat == KAT_MI else lembaga_dari_rating(rating))
        out.append({
            "Kategori": kat, "Nama": c["Nama"], "Rating": rating, "Lembaga": lembaga, "Outlook": outlook,
            "Tgl Rating": tgl_rating, "Terakhir Dicek": tgl_cek, "Hari Sejak Dicek": umur,
            "Rating Lembaga Lain": lain, "Tier": tier, "Bobot": bobot,
            "Limit Rumus": limit_rumus, "Limit Nominal": limit, "Sumber Limit": sumber, "Manual": manual,
            "Aktif": aktif, "Tanda": tanda, "Sorot": sorot,
            "Aktual Nominal": aktual, "Sisa Limit": sisa, "Utilisasi": util,
            "Sumber Rating": sumber_rating, "Catatan Cek": "; ".join(notes),
        })

    rows = pd.DataFrame(out, columns=[c for c in ROW_COLS if c != "No"])
    if not rows.empty:
        rows.insert(0, "No", rows.groupby("Kategori").cumcount() + 1)
    else:
        rows.insert(0, "No", pd.Series(dtype=int))
    res.rows = rows

    for k in set(rat_by_name) | set(ovr) | set(eks):
        if k not in known:
            res.warnings.append(f"'{k}' ada di data rating/status/eksposur tapi tidak ada di sheet Counterparty.")
    return res


# --------------------------------------------------------------------------
# Perbandingan dua periode
# --------------------------------------------------------------------------
KOLOM_BANDING = ["Kategori", "Nama", "Rating Lama", "Rating Baru", "Perubahan Rating",
                 "Limit Lama", "Limit Baru", "Selisih Limit"]


def bandingkan(sekarang: Hasil, pembanding: Hasil) -> pd.DataFrame:
    a = sekarang.rows.copy()
    b = pembanding.rows.copy()
    if a.empty and b.empty:
        return pd.DataFrame(columns=KOLOM_BANDING)
    a["_k"] = a["Nama"].map(norm_key)
    b["_k"] = b["Nama"].map(norm_key)
    cols = ["_k", "Kategori", "Nama", "Rating", "Limit Nominal"]
    mg = pd.merge(b[cols], a[cols], on="_k", how="outer", suffixes=(" Lama", " Baru"))
    int_labels = [l for l, _ in (sekarang.tiers_int or pembanding.tiers_int)]
    hasil = []
    for _, r in mg.iterrows():
        nama = r["Nama Baru"] if pd.notna(r["Nama Baru"]) else r["Nama Lama"]
        kat = r["Kategori Baru"] if pd.notna(r["Kategori Baru"]) else r["Kategori Lama"]
        rl, rb = r["Rating Lama"], r["Rating Baru"]
        ll, lb = r["Limit Nominal Lama"], r["Limit Nominal Baru"]
        if pd.isna(r["Nama Lama"]):
            ket = "Baru"
        elif pd.isna(r["Nama Baru"]):
            ket = "Keluar"
        else:
            labels = int_labels if kat == KAT_MI else None
            ra, rbk = rank_rating(rl, labels), rank_rating(rb, labels)
            if norm_key(rl) == norm_key(rb):
                ket = "Tetap"
            elif ra is not None and rbk is not None:
                ket = "Naik" if rbk < ra else ("Turun" if rbk > ra else "Ganti lembaga")
            else:
                ket = "Berubah"
        dl = ((lb if pd.notna(lb) else 0) - (ll if pd.notna(ll) else 0)) if (pd.notna(lb) or pd.notna(ll)) else None
        hasil.append({
            "Kategori": kat, "Nama": nama,
            "Rating Lama": rl if pd.notna(rl) else None, "Rating Baru": rb if pd.notna(rb) else None,
            "Perubahan Rating": ket,
            "Limit Lama": ll if pd.notna(ll) else None, "Limit Baru": lb if pd.notna(lb) else None,
            "Selisih Limit": dl,
        })
    df = pd.DataFrame(hasil, columns=KOLOM_BANDING)
    order = {n: i for i, n in enumerate(list(a["_k"]) + list(b["_k"]))}
    df["_o"] = df["Nama"].map(norm_key).map(order)
    df["_kat"] = (df["Kategori"] == KAT_MI).astype(int)
    return df.sort_values(["_kat", "_o"], kind="stable").drop(columns=["_o", "_kat"]).reset_index(drop=True)


def riwayat_bulanan(m: dict[str, pd.DataFrame], nama: str, akhir: date, n_bulan: int = 24,
                    mode: str = MODE_UTAMA) -> pd.DataFrame:
    """Rating & limit counterparty pada setiap akhir bulan (n bulan ke belakang)."""
    k = norm_key(nama)
    out = []
    y, mo = akhir.year, akhir.month
    titik = []
    for _ in range(n_bulan):
        titik.append(end_of_month(y, mo))
        mo -= 1
        if mo == 0:
            y, mo = y - 1, 12
    titik = sorted(set([t for t in titik if t <= akhir] + [akhir]))
    for t in titik:
        h = hitung(m, t, mode=mode, batas_umur_rating=0)
        if h.errors or h.rows.empty:
            continue
        r = h.rows[h.rows["Nama"].map(norm_key) == k]
        if r.empty:
            continue
        r = r.iloc[0]
        out.append({"Periode": pd.Timestamp(t), "Rating": r["Rating"], "Lembaga": r["Lembaga"],
                    "Limit Nominal": r["Limit Nominal"], "Total Aset Investasi": h.total_aset})
    return pd.DataFrame(out)

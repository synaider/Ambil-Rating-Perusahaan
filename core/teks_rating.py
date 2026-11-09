"""
Pembaca teks rilis rating (judul press release PEFINDO / Fitch) menjadi data terstruktur.

Contoh yang dikenali:
  "PEFINDO has raised the corporate rating of PT Waskita Karya (Persero) Tbk (WSKT) to idB from idSD"
  "PEFINDO has affirmed its idAAA ratings to PT Bank Rakyat Indonesia (Persero) Tbk (BBRI) and its outstanding bonds"
  "Bank Mega rated "idAA-" with stable outlook"
  "Fitch Affirms Bank Rakyat Indonesia at 'BBB'/'AAA(idn)'; Outlook Stable"
  "Fitch Upgrades Bank X to 'AA+(idn)'; Outlook Stable"
  "Fitch Rates X's Subordinated Bonds 'AA(idn)'"            -> instrumen (diabaikan untuk rating perusahaan)
"""
from __future__ import annotations

import difflib
import re
from dataclasses import dataclass
from datetime import date

import pandas as pd

from .engine import L_FITCH, L_PEFINDO, norm_key

# ----------------------------------------------------------------------------
# Token rating
# ----------------------------------------------------------------------------
# PEFINDO jangka panjang: idAAA, idAA+, ..., idSD, idD (+ akhiran (sy)/(gg)/(cg)/(bg)/(sf)).
# Skala jangka pendek (idA1+, idA1, idA2, ...) sengaja TIDAK cocok karena diikuti angka.
RE_PEFINDO = re.compile(
    r"(?<![A-Za-z0-9])id(AAA|AA|A|BBB|BB|B|CCC|SD|D)([+-]?)(?![A-Za-z0-9+])(\((?:sy|gg|cg|bg|sf|ir)\))?")
# Fitch skala nasional Indonesia: AAA(idn), AA+(idn), ... (F1+(idn) jangka pendek tidak cocok)
RE_FITCH = re.compile(r"(?<![A-Za-z0-9])(AAA|AA|A|BBB|BB|B|CCC|CC|C|RD|D)([+-]?)\s?\(idn\)")

RE_ACT = {
    "Naik": re.compile(r"\b(upgrad\w*|rais\w*|menaikkan|dinaikkan|naik)\b", re.I),
    "Turun": re.compile(r"\b(downgrad\w*|lower\w*|cut\w*|menurunkan|diturunkan|turun)\b", re.I),
    "Ditarik": re.compile(r"\b(withdr[ae]w\w*|withdrawn|mencabut|dicabut|ditarik)\b", re.I),
    "Afirmasi": re.compile(r"\b(affirm\w*|reaffirm\w*|menegaskan|mempertahankan|afirmasi)\b", re.I),
    "Watch": re.compile(r"\b(creditwatch|credit watch|rating watch|placed on|places)\b", re.I),
    "Baru": re.compile(r"\b(assign\w*|rated|rates|rate|publish\w*|memberikan|menetapkan|first[- ]time)\b", re.I),
}
RE_SKIP = re.compile(r"fully repaid|has repaid|telah melunasi|will mature|maturing in the near term|akan jatuh tempo|"
                     r"jatuh tempo dalam waktu dekat", re.I)
INSTR = (r"bonds?|obligasi|sukuk|mtns?|medium[- ]term|notes?|sbk|surat berharga|commercial papers?|subordinat\w*|"
         r"eba|kik|shelf[- ]?regist\w*|shelfregist\w*|phase|tahap|seri(?:es)?|tranche|programm?e?|instruments?|"
         r"debts?|securities|debentures?|senior|perpetual|loans?|ncd|negotiable")
RE_INSTR = re.compile(rf"\b({INSTR})\b", re.I)
RE_POSS_INSTR = re.compile(rf"(?:'|’)s\b[^.;]{{0,80}}?\b({INSTR})\b", re.I)
RE_CORP_STRONG = re.compile(r"corporate (credit )?ratings?|issuer (default )?ratings?|national long[- ]term rating|"
                            r"peringkat (perusahaan|korporasi)|\band its\b|\bdan (obligasi|sukuk)\w*", re.I)
RE_AFTER_TOKEN_INSTR = re.compile(
    rf"^\W*(?:ratings?\s+)?(?:for|to|of|on|untuk|atas)\s+(?:the\s+|its\s+)?(?:proposed\s+|rencana\s+)?"
    rf"(?:\S+\s+){{0,3}}?({INSTR})\b", re.I)
RE_OUTLOOK = re.compile(
    r"(?:(stable|negative|positive|developing|stabil|negatif|positif|berkembang)[\"']?\s+outlook)|"
    r"(?:outlook\s*(?:is\s+|to\s+|revised to\s+)?(stable|negative|positive|developing|stabil|negatif|positif|berkembang))|"
    r"(?:(?:credit|rating)\s?watch\s+(?:with\s+)?(negative|positive|developing|evolving)(?:\s+implications?)?)|"
    r"(?:\bRW([NPE])\b)", re.I)
_OUTLOOK_MAP = {"stabil": "Stable", "negatif": "Negative", "positif": "Positive", "berkembang": "Developing",
                "evolving": "Developing", "n": "Negative", "p": "Positive", "e": "Developing"}


@dataclass
class Token:
    lembaga: str
    rating: str          # teks persis, mis. 'idAA+', 'idAAA(gg)', 'AA(idn)'
    start: int
    end: int


def cari_token(teks: str) -> list[Token]:
    toks = []
    for m in RE_PEFINDO.finditer(teks):
        toks.append(Token(L_PEFINDO, "id" + m.group(1) + (m.group(2) or "") + (m.group(3) or ""), m.start(), m.end()))
    for m in RE_FITCH.finditer(teks):
        toks.append(Token(L_FITCH, m.group(1) + (m.group(2) or "") + "(idn)", m.start(), m.end()))
    return sorted(toks, key=lambda t: t.start)


def _rapikan(teks: str) -> str:
    t = str(teks or "")
    t = t.replace("“", '"').replace("”", '"').replace("‘", "'").replace("’", "'")
    return re.sub(r"\s+", " ", t).strip()


def baca_outlook(teks: str) -> str | None:
    m = RE_OUTLOOK.search(teks)
    if not m:
        return None
    if m.group(3):
        return "CreditWatch " + _OUTLOOK_MAP.get(m.group(3).lower(), m.group(3).capitalize())
    if m.group(4):
        return "Rating Watch " + _OUTLOOK_MAP.get(m.group(4).lower(), m.group(4))
    w = (m.group(1) or m.group(2)).lower()
    return _OUTLOOK_MAP.get(w, w.capitalize())


def baca_aksi(teks: str) -> str:
    for nama in ("Naik", "Turun", "Ditarik", "Afirmasi", "Watch", "Baru"):
        if RE_ACT[nama].search(teks):
            return nama
    return ""


def tingkat(teks: str, tokens: list[Token]) -> str:
    """'perusahaan' | 'instrumen' | 'lain' (tidak ada aksi rating)."""
    aksi = baca_aksi(teks)
    if RE_SKIP.search(teks) and aksi not in ("Naik", "Turun", "Afirmasi", "Ditarik"):
        return "lain"
    if not aksi and not tokens:
        return "lain"
    if RE_CORP_STRONG.search(teks):
        return "perusahaan"
    if RE_POSS_INSTR.search(teks):
        return "instrumen"
    if tokens:
        head = teks[:tokens[0].start]
        tail = teks[tokens[0].end:tokens[0].end + 80]
    else:
        head, tail = teks, ""
    if RE_INSTR.search(head):
        return "instrumen"
    if RE_AFTER_TOKEN_INSTR.search(tail):
        return "instrumen"
    return "perusahaan"


def pilih_token(teks: str, tokens: list[Token], lembaga: str | None = None) -> Token | None:
    """Rating baru dari kalimat aksi: 'to X from Y' -> X; 'from Y to X' -> X; 'dari Y menjadi X' -> X;
    selain itu token pertama (lembaga yang diminta, bila ada)."""
    tok = [t for t in tokens if lembaga is None or t.lembaga == lembaga] or tokens
    if not tok:
        return None
    for t in tok:
        before = teks[max(0, t.start - 14):t.start].lower()
        if re.search(r"\b(to|menjadi|ke)\s*['\"]?\s*$", before):
            return t
    return tok[0]


def baca_judul(judul: str, lembaga_hint: str | None = None) -> dict:
    """Analisis satu judul/kalimat rilis."""
    t = _rapikan(judul)
    tokens = cari_token(t)
    lembaga = lembaga_hint
    if lembaga is None:
        if tokens:
            lembaga = tokens[0].lembaga
        elif re.search(r"\bfitch\b", t, re.I):
            lembaga = L_FITCH
        elif re.search(r"\bpefindo\b", t, re.I):
            lembaga = L_PEFINDO
    lvl = tingkat(t, tokens)
    tok = pilih_token(t, tokens, lembaga)
    aksi = baca_aksi(t)
    rating = tok.rating if tok else None
    if aksi == "Ditarik" and lvl == "perusahaan" and rating is None:
        rating = "NR"
    return {"teks": t, "lembaga": lembaga, "tingkat": lvl, "aksi": aksi, "rating": rating,
            "outlook": baca_outlook(t)}


# ----------------------------------------------------------------------------
# Tanggal
# ----------------------------------------------------------------------------
_BULAN = {
    "jan": 1, "januari": 1, "january": 1, "feb": 2, "februari": 2, "february": 2, "peb": 2, "mar": 3, "maret": 3,
    "march": 3, "apr": 4, "april": 4, "mei": 5, "may": 5, "jun": 6, "juni": 6, "june": 6, "jul": 7, "juli": 7,
    "july": 7, "agu": 8, "agt": 8, "ags": 8, "agustus": 8, "aug": 8, "august": 8, "sep": 9, "sept": 9,
    "september": 9, "okt": 10, "oktober": 10, "oct": 10, "october": 10, "nov": 11, "nop": 11, "november": 11,
    "nopember": 11, "des": 12, "desember": 12, "dec": 12, "december": 12,
}
_BLN_RE = "|".join(sorted(_BULAN, key=len, reverse=True))
RE_TGL = [
    (re.compile(r"\b(\d{4})-(\d{1,2})-(\d{1,2})\b"), "ymd"),
    (re.compile(r"\b(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})\b"), "dmy"),
    (re.compile(rf"\b(\d{{1,2}})\s+({_BLN_RE})\.?,?\s+(\d{{4}})\b", re.I), "d_mon_y"),
    (re.compile(rf"\b({_BLN_RE})\.?\s+(\d{{1,2}}),?\s+(\d{{4}})\b", re.I), "mon_d_y"),
]


def cari_tanggal(teks: str) -> tuple[date | None, tuple[int, int] | None]:
    """Cari tanggal pertama di teks. Return (tanggal, (start, end))."""
    best = None
    for rx, fmt in RE_TGL:
        for m in rx.finditer(teks or ""):
            try:
                if fmt == "ymd":
                    d = date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
                elif fmt == "dmy":
                    d = date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
                elif fmt == "d_mon_y":
                    d = date(int(m.group(3)), _BULAN[m.group(2).lower()], int(m.group(1)))
                else:
                    d = date(int(m.group(3)), _BULAN[m.group(1).lower()], int(m.group(2)))
            except (ValueError, KeyError):
                continue
            if best is None or m.start() < best[1][0]:
                best = (d, (m.start(), m.end()))
            break
    return best if best else (None, None)


def parse_tanggal(s) -> date | None:
    if isinstance(s, date):
        return s
    d, _ = cari_tanggal(str(s or ""))
    return d


# ----------------------------------------------------------------------------
# Pencocokan ke counterparty
# ----------------------------------------------------------------------------
_STOP = re.compile(r"\b(pt|tbk|persero|perseroda|the|bank|group|grup|limited|ltd|plc|company|co)\b\.?", re.I)


def _norm_nama(s: str) -> str:
    s = re.sub(r"[^\w\s]", " ", str(s or "").lower())
    return re.sub(r"\s+", " ", s).strip()


def alias_counterparty(cp: pd.DataFrame) -> list[tuple[str, str]]:
    """List (alias_ternormalisasi, nama_counterparty)."""
    out = []
    for _, r in cp.iterrows():
        cands = [r.get("Nama Entitas"), r.get("Kode PEFINDO")] + str(r.get("Alias") or "").split(";")
        nm = re.sub(r"[+\-]\)\s*$", "", str(r["Nama"]))
        nm_inti = re.sub(r"\b(group|grup)\b", "", nm, flags=re.I)
        cands += [nm, nm_inti]
        seen = set()
        for c in cands:
            if not c or (isinstance(c, float) and pd.isna(c)):
                continue
            a = _norm_nama(c)
            a2 = _norm_nama(_STOP.sub(" ", a))
            for x in (a, a2):
                if len(x) >= 3 and x not in seen:
                    seen.add(x)
                    out.append((x, r["Nama"]))
    return out


def cocokkan(teks: str, aliases: list[tuple[str, str]]) -> tuple[str | None, float]:
    """Cari counterparty yang aliasnya muncul di teks (alias terpanjang menang)."""
    t = " " + _norm_nama(teks) + " "
    best, best_len = None, 0
    for a, nama in aliases:
        if f" {a} " in t and len(a) > best_len:
            best, best_len = nama, len(a)
    if best:
        return best, 1.0
    # fuzzy fallback (judul berisi nama entitas yang sedikit berbeda)
    words = t.split()
    best_score = 0.0
    for a, nama in aliases:
        n = len(a.split())
        if n < 2:
            continue
        for i in range(0, max(1, len(words) - n + 1)):
            frag = " ".join(words[i:i + n])
            sc = difflib.SequenceMatcher(None, a, frag).ratio()
            if sc > best_score:
                best, best_score = nama, sc
    return (best, best_score) if best_score >= 0.86 else (None, best_score)


# ----------------------------------------------------------------------------
# Teks tempel bebas -> tabel rating
# ----------------------------------------------------------------------------
KOLOM_PARSE = ["Simpan", "Tanggal", "Nama", "Lembaga", "Rating", "Outlook", "Aksi", "Tingkat", "Teks"]


def parse_teks_bebas(teks: str, cp: pd.DataFrame, tanggal_default: date | None = None,
                     lembaga_hint: str | None = None) -> pd.DataFrame:
    """Teks hasil copy-paste (judul rilis per baris, boleh diawali/diakhiri tanggal; atau format
    'tanggal' di satu baris lalu 'judul' di baris berikutnya)."""
    aliases = alias_counterparty(cp)
    baris = [b.strip() for b in re.split(r"[\r\n]+", teks or "") if b.strip()]
    out = []
    tgl_pending = None
    for b in baris:
        d, span = cari_tanggal(b)
        sisa = b
        if span:
            sisa = (b[:span[0]] + " " + b[span[1]:]).strip(" \t|;,-–")
        if d and len(re.sub(r"\W", "", sisa)) < 6:     # baris berisi tanggal saja
            tgl_pending = d
            continue
        info = baca_judul(sisa, lembaga_hint)
        if info["tingkat"] == "lain" and not info["rating"]:
            tgl_pending = None
            continue
        tgl = d or tgl_pending or tanggal_default
        tgl_pending = None
        nama, skor = cocokkan(sisa, aliases)
        out.append({
            "Simpan": bool(info["tingkat"] == "perusahaan" and info["rating"] and nama and tgl),
            "Tanggal": pd.Timestamp(tgl) if tgl else pd.NaT,
            "Nama": nama, "Lembaga": info["lembaga"], "Rating": info["rating"], "Outlook": info["outlook"],
            "Aksi": info["aksi"], "Tingkat": info["tingkat"], "Teks": info["teks"],
        })
    return pd.DataFrame(out, columns=KOLOM_PARSE)

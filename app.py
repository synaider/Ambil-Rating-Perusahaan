"""
Cek Rating & Update Limit Counterparty (Streamlit)

Alur: pilih tanggal → cek rating pada tanggal itu (PEFINDO → Fitch → perusahaan anggota group)
      → download file Excel limit Anda dengan rating & tanggal yang sudah diganti.

Jalankan:  streamlit run app.py
"""
from __future__ import annotations

import hashlib
import inspect
import re
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

from core import pefindo
from core.defaults import key_group
from core.engine import L_FITCH, parse_rating
from core.pefindo import riwayat_dari_emiten
from core.sumber import (KOLOM_FITCH, LEMBAGA_PILIHAN, SumberPefindo, cek_group, fitch_dari_template, keterangan, muat_pengaturan,
                         rapikan_entitas, rapikan_fitch, rating_fitch_pada, rating_pefindo_pada, simpan_pengaturan,
                         url_fitch, url_pefindo)
from core.teks_rating import parse_teks_bebas
from core.template_excel import baca_template, hitung_limit, tulis_template

APP_DIR = Path(__file__).resolve().parent
DATA = APP_DIR / "data"
CONTOH = DATA / "contoh" / "Limit_Countreparty_09102026.xlsx"
TPL_FILE = DATA / "template_terakhir.xlsx"
TPL_NAMA = DATA / "template_terakhir.txt"
PENGATURAN = DATA / "pengaturan.xlsx"
CACHE = DATA / "cache_pefindo"
BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober",
         "November", "Desember"]

st.set_page_config(page_title="Cek Rating Counterparty", page_icon="📊", layout="wide",
                   initial_sidebar_state="collapsed")


# ----------------------------------------------------------------------------
# Kompatibilitas versi Streamlit (parameter yang belum dikenal versi lama diabaikan)
# ----------------------------------------------------------------------------
def _params(fn) -> set:
    try:
        return set(inspect.signature(fn).parameters)
    except (TypeError, ValueError):
        return set()


def kol(jenis: str, **kw):
    fn = getattr(st.column_config, jenis)
    p = _params(fn)
    return fn(**{k: v for k, v in kw.items() if k in p})


LEBAR = {"width": "stretch"} if "width" in _params(st.download_button) else {"use_container_width": True}


def tgl_id(d) -> str:
    if d is None or (not isinstance(d, date) and pd.isna(d)):
        return "—"
    d = pd.Timestamp(d)
    return f"{d.day} {BULAN[d.month - 1]} {d.year}"


def angka(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "NA"
    return f"{v:,.0f}".replace(",", ".")


# ----------------------------------------------------------------------------
# Data: file Excel template & pengaturan sumber
# ----------------------------------------------------------------------------
def template() -> tuple[bytes, str]:
    if TPL_FILE.exists():
        nama = TPL_NAMA.read_text(encoding="utf-8").strip() if TPL_NAMA.exists() else TPL_FILE.name
        return TPL_FILE.read_bytes(), nama
    return CONTOH.read_bytes(), CONTOH.name


def simpan_template(data: bytes, nama: str):
    DATA.mkdir(parents=True, exist_ok=True)
    TPL_FILE.write_bytes(data)
    TPL_NAMA.write_text(nama, encoding="utf-8")


def pengaturan() -> tuple[pd.DataFrame, pd.DataFrame]:
    if "ent" not in st.session_state:
        baru = not PENGATURAN.exists()
        ent, fit, disesuaikan = muat_pengaturan(PENGATURAN)
        if disesuaikan:   # pengaturan versi lama → simpan versi yang sudah disesuaikan
            simpan_pengaturan(PENGATURAN, ent, fit)
        if baru:     # data Fitch awal = rating (idn) yang tertulis di file contoh
            data, nama = template()
            s = baca_template(data, nama)
            fit = rapikan_fitch(fitch_dari_template(s["nonmi"], s["tanggal"], ent, nama))
            simpan_pengaturan(PENGATURAN, ent, fit)
        st.session_state.ent, st.session_state.fit = ent, fit
    return st.session_state.ent, st.session_state.fit


def simpan_set(ent=None, fit=None, pesan="Tersimpan."):
    e = rapikan_entitas(ent if ent is not None else st.session_state.ent)
    f = rapikan_fitch(fit if fit is not None else st.session_state.fit)
    simpan_pengaturan(PENGATURAN, e, f)
    st.session_state.ent, st.session_state.fit = e, f
    st.session_state.pop("hasil", None)
    st.session_state.flash = pesan


@st.cache_data(ttl=6 * 3600, show_spinner=False)
def daftar_pefindo() -> pd.DataFrame:
    return pefindo.ambil_daftar_emiten()


ent, fit = pengaturan()
if st.session_state.get("flash"):
    st.success(st.session_state.pop("flash"))

st.title("Cek Rating & Update Limit Counterparty")
st.caption("Pilih tanggal → cek rating pada tanggal itu (PEFINDO → Fitch → perusahaan anggota group) → "
           "download file Excel limit dengan rating yang sudah diperbarui.")

tab_cek, tab_cari, tab_sumber, tab_panduan = st.tabs(
    ["📋 Cek & Update Rating", "🔎 Cari Rating", "⚙️ Sumber Rating", "❓ Panduan"])

# ============================================================================
# TAB 1
# ============================================================================
with tab_cek:
    c1, c2 = st.columns([1, 2])
    tgl = c1.date_input("1️⃣ Tanggal periode", value=date.today(), format="DD/MM/YYYY",
                        help="Rating yang diambil = rating terakhir yang terbit PADA atau SEBELUM tanggal ini.")
    up = c2.file_uploader("2️⃣ File Excel limit (kosongkan = pakai file terakhir)", type=["xlsx", "xlsm"])
    if up is not None:
        data_up = up.getvalue()
        try:
            baca_template(data_up, up.name)
            if hashlib.md5(data_up).hexdigest() != st.session_state.get("tpl_md5_up"):
                simpan_template(data_up, up.name)
                st.session_state.tpl_md5_up = hashlib.md5(data_up).hexdigest()
                st.session_state.pop("hasil", None)
        except Exception as e:  # noqa: BLE001
            st.error(f"File tidak bisa dibaca sebagai file limit counterparty: {e}")
    data_tpl, nama_tpl = template()
    snap = baca_template(data_tpl, nama_tpl)
    kunci_tpl = hashlib.md5(data_tpl).hexdigest()
    c2.caption(f"File dipakai: **{nama_tpl}** · tanggal di file {tgl_id(snap['tanggal'])} · "
               f"{len(snap['nonmi'])} counterparty Non MI")

    with st.expander("Total Aset Investasi (opsional — biarkan bila tidak berubah)"):
        e1, e2 = st.columns(2)
        ta = e1.number_input("Total Aset Investasi (miliar Rp)", min_value=0.0,
                             value=float(snap["total_aset"] or 0.0), step=1.0, format="%.2f")
        tgl_ta = e2.date_input("Posisi tanggal", value=snap["total_aset_tgl"] or tgl, format="DD/MM/YYYY")
    ta_berubah = abs(ta - float(snap["total_aset"] or 0.0)) > 1e-9 or tgl_ta != snap["total_aset_tgl"]

    b1, b2 = st.columns([1, 2])
    cek = b1.button(f"3️⃣ 🔍 Cek rating per {tgl:%d/%m/%Y}", type="primary", **LEBAR)
    paksa = b2.checkbox("Ambil ulang dari PEFINDO (abaikan data yang sudah diambil hari ini)")

    if cek:
        pef = SumberPefindo(CACHE, paksa_ulang=paksa)
        rows = []
        bar = st.progress(0.0, text="Memulai…")
        n = len(snap["nonmi"])
        for i, r in enumerate(snap["nonmi"], start=1):
            bar.progress(i / n, text=f"{i}/{n} · {r['Nama']}")
            tm = cek_group(r["Nama"], ent, fit, tgl, pef)
            baru = tm.rating or r["Rating"]
            rows.append({"No": i, "Counterparty": r["Nama"], "Rating di file": r["Rating"], "Rating periode ini": baru,
                         "Berubah": "✔" if (r["Rating"] or "") != (baru or "") else "",
                         "Lembaga": tm.lembaga or "—", "Link": tm.url, "Entitas sumber": tm.entitas or "—",
                         "Tgl rating": tm.tanggal.strftime("%d/%m/%Y") if tm.tanggal else "",
                         "Outlook": tm.outlook or "", "Keterangan": keterangan(tm), "_ok": tm.rating is not None})
        bar.empty()
        st.session_state.hasil = {"tgl": tgl, "tpl": kunci_tpl, "df": pd.DataFrame(rows),
                                  "offline": pef.gagal_koneksi >= 2 and pef.sukses == 0}

    hasil = st.session_state.get("hasil")
    if hasil and (hasil["tgl"] != tgl or hasil["tpl"] != kunci_tpl):
        st.info("Tanggal atau file berubah — klik **Cek rating** lagi.")
    elif hasil:
        df = hasil["df"]
        if hasil["offline"]:
            st.warning("pefindo.com tidak bisa diakses dari komputer ini (cek internet / jaringan kantor). "
                       "Rating PEFINDO memakai data tersimpan terakhir bila ada; sisanya tetap memakai rating di file.")
        ok = df["_ok"]
        n_pef = int((df["Lembaga"] == "PEFINDO").sum())
        n_fit = int((df["Lembaga"] == L_FITCH).sum())
        n_ang = int(df["Keterangan"].str.contains("anggota group").sum())
        n_tdk = int((~ok).sum())
        n_ubah = int((df["Berubah"] == "✔").sum())
        st.markdown(f"**Rating per {tgl_id(tgl)}** — PEFINDO **{n_pef}** · Fitch **{n_fit}** · "
                    f"dari anggota group **{n_ang}** · tidak ditemukan **{n_tdk}** · berubah dari file **{n_ubah}**")
        st.caption("Kolom *Rating periode ini* bisa diedit sebelum download. Baris yang tidak ditemukan tetap memakai "
                   "rating di file.")
        tampil = df.drop(columns="_ok")
        ed = st.data_editor(
            tampil, hide_index=True, key=f"ed_{hasil['tgl']}_{hasil['tpl']}",
            disabled=[c for c in tampil.columns if c != "Rating periode ini"],
            height=min(40 + 35 * len(tampil), 760),
            column_config={
                "No": kol("NumberColumn", width="small"),
                "Counterparty": kol("TextColumn", width="medium"),
                "Rating di file": kol("TextColumn", width="medium"),
                "Rating periode ini": kol("TextColumn", width="medium", help="Bisa diedit"),
                "Berubah": kol("TextColumn", width="small"),
                "Lembaga": kol("TextColumn", width="small"),
                "Link": kol("LinkColumn", width="small", display_text=r"https?://(?:www\.)?([a-z]+)\.",
                            help="Halaman sumber: PEFINDO (halaman emiten) atau pencarian di situs Fitch"),
                "Entitas sumber": kol("TextColumn", width="medium"),
                "Tgl rating": kol("TextColumn", width="small"),
                "Keterangan": kol("TextColumn", width="large"),
            })
        rating_akhir = {r["Nama"]: (v.strip() if isinstance(v, str) else v)
                        for r, v in zip(snap["nonmi"], ed["Rating periode ini"])}
        salah = [k for k, v in rating_akhir.items() if v and not parse_rating(v)]
        if salah:
            st.warning("Format rating tidak dikenali (limit akan NA): " + ", ".join(salah))
        lim_baru = hitung_limit(snap, rating_akhir, ta)
        lim_lama = hitung_limit(snap, {r["Nama"]: r["Rating"] for r in snap["nonmi"]})
        tot_baru = sum(v or 0 for v in lim_baru.values())
        tot_lama = sum(v or 0 for v in lim_lama.values())
        m1, m2, m3 = st.columns(3)
        m1.metric("Total limit Non MI (periode ini)", angka(tot_baru), delta=angka(tot_baru - tot_lama)
                  if tot_baru != tot_lama else None)
        m2.metric("Total limit Non MI (file)", angka(tot_lama))
        m3.metric("Total Aset Investasi", f"{ta:,.2f}".replace(",", "_").replace(".", ",").replace("_", "."))

        perbaiki = st.checkbox("Perbaiki rumus total agar menjumlah semua baris counterparty", value=True,
                               help="Di file contoh, TOTAL NON MANAJER INVESTASI memakai SUM(D7:D47) sehingga 3 baris "
                                    "terakhir tidak ikut terjumlah.")
        xls = tulis_template(data_tpl, snap, tgl, rating_akhir,
                             total_aset=ta if ta_berubah else None, tgl_aset=tgl_ta if ta_berubah else None,
                             perbaiki_total=perbaiki)
        st.download_button(f"4️⃣ ⬇️ Download Excel — Limit_Countreparty_{tgl:%d%m%Y}.xlsx", data=xls,
                           file_name=f"Limit_Countreparty_{tgl:%d%m%Y}.xlsx", type="primary",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", **LEBAR)
        st.caption("Yang diganti di file: tanggal laporan, kolom Rating (Non MI)" +
                   (", Total Aset Investasi" if ta_berubah else "") +
                   ". Format, rumus limit, dan bagian Manajer Investasi tetap seperti file asli; "
                   "limit otomatis dihitung ulang saat file dibuka di Excel.")

# ============================================================================
# TAB 2 – CARI RATING (perusahaan di luar daftar default)
# ============================================================================
with tab_cari:
    st.markdown("#### Cari rating perusahaan lain")
    st.caption("Untuk perusahaan yang tidak ada di daftar default: aplikasi memindai daftar emiten PEFINDO lalu "
               "menampilkan rating pada tanggal yang dipilih beserta riwayatnya. Untuk Fitch disediakan link "
               "pencarian langsung ke situs Fitch.")
    with st.form("form_cari"):
        f1, f2 = st.columns([3, 1])
        q_in = f1.text_input("Nama perusahaan atau kode PEFINDO", placeholder="mis. Indofood CBP, Pupuk, BTPN")
        tgl_c_in = f2.date_input("Rating per tanggal", value=tgl, format="DD/MM/YYYY")
        scan = st.form_submit_button("🔍 Scan PEFINDO & Fitch", type="primary")
    if scan and q_in.strip():
        q0 = q_in.strip()
        err = None
        try:
            hasil_c = pefindo.cari_kode(daftar_pefindo(), q0, n=10)
            hasil_c = hasil_c[hasil_c["Skor"] >= 0.45].drop(columns="Skor")
        except Exception as e:  # noqa: BLE001
            hasil_c = pd.DataFrame(columns=["Nama Perusahaan", "Kode", "Jenis Rating", "Industri"])
            err = str(e)
        if re.fullmatch(r"[A-Za-z0-9]{3,6}", q0) and q0.upper() not in set(hasil_c["Kode"]):
            hasil_c = pd.concat([pd.DataFrame([{"Nama Perusahaan": f"(kode {q0.upper()})", "Kode": q0.upper()}]),
                                 hasil_c], ignore_index=True)
        hasil_c["Link"] = hasil_c["Kode"].map(url_pefindo)
        st.session_state.cari = {"q": q0, "tgl": tgl_c_in, "hasil": hasil_c.reset_index(drop=True), "err": err,
                                 "data": None}
    cr = st.session_state.get("cari")
    if cr:
        q0, tgl_c = cr["q"], cr["tgl"]
        kiri, kanan = st.columns([3, 2], gap="large")
        with kiri:
            st.markdown(f"##### PEFINDO — “{q0}”")
            if cr["err"]:
                st.warning(f"Daftar emiten PEFINDO tidak bisa diambil ({cr['err'][:120]}). "
                           "Jika tahu kodenya, ketik kode PEFINDO langsung (mis. ICBP).")
            hc = cr["hasil"]
            if hc.empty:
                st.info("Tidak ditemukan di daftar emiten PEFINDO.")
            else:
                st.dataframe(hc.fillna(""), hide_index=True,
                             column_config={"Link": kol("LinkColumn", display_text="Buka PEFINDO", width="small")})
                p1, p2 = st.columns([3, 1])
                nama_kode = dict(zip(hc["Kode"], hc["Nama Perusahaan"]))
                kode_pil = p1.selectbox("Pilih perusahaan", list(nama_kode),
                                        format_func=lambda k: f"{k} — {nama_kode.get(k, '')}")
                p2.write("")
                if p2.button("Ambil rating", type="primary", **LEBAR):
                    with st.spinner(f"Mengambil {kode_pil} dari pefindo.com…"):
                        d, cat = SumberPefindo(CACHE).ambil(kode_pil)
                    cr["data"], cr["cat"], cr["kode"] = d, cat, kode_pil
            d = cr.get("data")
            if cr.get("kode"):
                kd = cr["kode"]
                if d is None:
                    st.error(f"Data {kd} tidak bisa diambil: {cr.get('cat')}")
                else:
                    r, why = rating_pefindo_pada(d, tgl_c)
                    nm = d.nama or kd
                    st.markdown(f"**{nm}** ({kd})")
                    m1, m2 = st.columns(2)
                    m1.metric(f"Rating per {tgl_c:%d/%m/%Y}", r["Rating"] if r is not None else "—",
                              help=None if r is not None else why)
                    m2.metric("Rating terkini di PEFINDO", f"{d.rating or '—'} {d.outlook or ''}".strip(),
                              help=f"Updated on {d.tgl_update:%d/%m/%Y}" if d.tgl_update else None)
                    if r is not None:
                        st.caption(f"Dari rilis {r['Tanggal Efektif']:%d/%m/%Y}" +
                                   (f" · outlook {r['Outlook']}" if r["Outlook"] else ""))
                    else:
                        st.caption(f"Tidak ada rating berlaku per {tgl_c:%d/%m/%Y}: {why}")
                    if cr.get("cat"):
                        st.caption(cr["cat"])
                    st.link_button(f"🔗 Buka halaman {kd} di PEFINDO", d.url)
                    rw, _ = riwayat_dari_emiten(d)
                    if rw:
                        rwd = pd.DataFrame(rw)[["Tanggal Efektif", "Rating", "Outlook", "Keterangan"]].iloc[::-1]
                        rwd["Tanggal Efektif"] = rwd["Tanggal Efektif"].dt.strftime("%d/%m/%Y")
                        st.dataframe(rwd.fillna(""), hide_index=True, height=min(40 + 35 * len(rwd), 320),
                                     column_config={"Keterangan": kol("TextColumn", width="large")})
                    with st.form("tambah_ke_group"):
                        st.markdown("**Tambahkan ke daftar entitas group** (supaya ikut dicek otomatis)")
                        g1, g2 = st.columns(2)
                        grup_ada = sorted(ent["Group"].unique())
                        g_pil = g1.selectbox("Group", ["(group baru)"] + grup_ada)
                        g_baru = g2.text_input("Nama group baru (jika group baru)", placeholder="mis. INDOFOOD GROUP")
                        posisi = st.radio("Posisi", ["Anggota group (urutan terakhir)", "Entitas utama (urutan 1)"],
                                          horizontal=True)
                        if st.form_submit_button("➕ Tambahkan"):
                            grup = key_group(g_baru) if g_pil == "(group baru)" else g_pil
                            if not grup:
                                st.error("Isi nama group baru.")
                            else:
                                urut = 0 if posisi.startswith("Entitas utama") else 999
                                e_baru = pd.concat([ent, pd.DataFrame([{
                                    "Group": grup, "Urutan": urut, "Entitas": nm, "Kode PEFINDO": kd,
                                    "Lembaga": "Otomatis", "Alias": None}])], ignore_index=True)
                                simpan_set(ent=e_baru, pesan=f"{nm} ({kd}) ditambahkan ke {grup}.")
                                st.rerun()
        with kanan:
            st.markdown(f"##### Fitch Ratings — “{q0}”")
            st.link_button(f"🔗 Cari “{q0}” di situs Fitch", url_fitch(q0))
            st.caption("Situs Fitch tidak bisa dipindai otomatis. Buka link di atas, lalu catat ratingnya di bawah "
                       "(atau tempel judul rilisnya di tab Sumber Rating).")
            cocok = fit[fit["Entitas"].str.contains(re.escape(q0), case=False, na=False)]
            if not cocok.empty:
                st.markdown("Sudah ada di tabel Fitch:")
                for nm_f in cocok["Entitas"].unique():
                    rf, whyf = rating_fitch_pada(fit, nm_f, tgl_c)
                    st.write(f"• {nm_f}: **{rf['Rating'] if rf is not None else '—'}** per {tgl_c:%d/%m/%Y}"
                             + ("" if rf is not None else f" ({whyf})"))
            with st.form("tambah_fitch_cepat"):
                st.markdown("**Catat rating Fitch**")
                fa, fb, fc = st.columns([3, 1.4, 1.2])
                f_ent = fa.text_input("Entitas", value=q0)
                f_tgl = fb.date_input("Tanggal rilis", value=tgl_c, format="DD/MM/YYYY")
                f_rat = fc.text_input("Rating", placeholder="AA+(idn)")
                f_ket = st.text_input("Keterangan / judul rilis", placeholder="Fitch Affirms ... ; Outlook Stable")
                if st.form_submit_button("➕ Simpan ke tabel Fitch"):
                    rr = f_rat.strip()
                    if not rr.endswith("(idn)") and parse_rating(rr + "(idn)"):
                        rr += "(idn)"
                    if not parse_rating(rr):
                        st.error("Format rating tidak dikenali. Contoh: AAA(idn), AA-(idn).")
                    else:
                        f_baru = pd.concat([fit, pd.DataFrame([{"Entitas": f_ent.strip(), "Tanggal": pd.Timestamp(f_tgl),
                                                               "Rating": rr, "Outlook": None,
                                                               "Keterangan": f_ket or "Input dari tab Cari Rating"}])],
                                           ignore_index=True)
                        simpan_set(fit=f_baru, pesan=f"Rating Fitch {rr} untuk {f_ent} tersimpan. Agar ikut dicek "
                                                      f"otomatis, pastikan nama entitas ini ada di daftar Entitas group.")
                        st.rerun()

# ============================================================================
# TAB 3 – SUMBER RATING
# ============================================================================
with tab_sumber:
    st.markdown("#### Entitas per group")
    st.caption("Urutan 1 = entitas utama. Kolom Lembaga: Otomatis = cek PEFINDO dulu lalu Fitch; PEFINDO / Fitch = "
               "hanya lembaga itu. Kalau tidak ada, lanjut ke urutan berikutnya (perusahaan anggota group). "
               "Kode PEFINDO = kode di pefindo.com.")
    belum = sorted({key_group(r["Nama"]) for r in snap["nonmi"]} - set(ent["Group"]))
    if belum:
        st.warning("Group di file yang belum punya entitas: " + ", ".join(belum))
    ed_ent = st.data_editor(ent, num_rows="dynamic", hide_index=True, key="ed_ent",
                            height=min(40 + 35 * (len(ent) + 2), 520),
                            column_config={"Group": kol("TextColumn", width="medium"),
                                           "Urutan": kol("NumberColumn", width="small", step=1),
                                           "Entitas": kol("TextColumn", width="large"),
                                           "Kode PEFINDO": kol("TextColumn", width="small"),
                                           "Lembaga": kol("SelectboxColumn", options=LEMBAGA_PILIHAN, width="small",
                                                          help="Otomatis = PEFINDO dulu lalu Fitch; PEFINDO/Fitch = "
                                                               "hanya lembaga itu"),
                                           "Alias": kol("TextColumn", width="large",
                                                        help="Nama lain untuk mengenali teks rilis Fitch (pisah ;)")})
    if st.button("💾 Simpan entitas group"):
        simpan_set(ent=ed_ent, pesan="Entitas group tersimpan.")
        st.rerun()

    with st.expander("🔎 Cari Kode PEFINDO"):
        with st.form("cari_kode"):
            q = st.text_input("Nama perusahaan / kode", placeholder="mis. Indofood, Sinar Mas, Pegadaian")
            cari = st.form_submit_button("Cari")
        if cari and q:
            try:
                h = pefindo.cari_kode(daftar_pefindo(), q)
                if h.empty:
                    st.info("Tidak ditemukan.")
                else:
                    st.dataframe(h.drop(columns="Skor"), hide_index=True)
            except Exception as e:  # noqa: BLE001
                st.error(f"Gagal mengambil daftar emiten dari pefindo.com: {e}")

    st.divider()
    st.markdown("#### Rating Fitch")
    st.info("Rating Fitch tidak bisa diambil otomatis (situs Fitch tidak menyediakan akses data terbuka). "
            "Saat Fitch menerbitkan rating baru, tempel judul rilisnya di bawah — tanggal, rating, dan entitas "
            "dibaca otomatis. Aplikasi memakai rating Fitch terakhir yang tanggalnya ≤ tanggal periode.")
    with st.form("form_fitch"):
        teks = st.text_area("Judul rilis Fitch (satu per baris, dengan tanggal)", height=120, placeholder=(
            "12 Mar 2026  Fitch Affirms Bank OCBC NISP at 'AAA(idn)'; Outlook Stable\n"
            "2026-05-20 | Fitch Upgrades Bank SulutGo to 'A+(idn)'; Outlook Stable"))
        baca = st.form_submit_button("Baca teks")
    if baca and teks.strip():
        cp = pd.DataFrame({"Nama": ent["Entitas"], "Nama Entitas": ent["Entitas"], "Alias": ent["Alias"],
                           "Kode PEFINDO": ent["Kode PEFINDO"]}).drop_duplicates("Nama")
        p = parse_teks_bebas(teks, cp, tgl, L_FITCH)
        p = p[p["Rating"].fillna("").str.contains(r"\(idn\)", regex=True)]
        st.session_state.parsed_fitch = p
    pf = st.session_state.get("parsed_fitch")
    if pf is not None:
        if pf.empty:
            st.warning("Tidak ada rating Fitch (format '...(idn)') yang terbaca.")
        else:
            st.caption("Periksa hasil baca. Baris 'instrumen' (obligasi/sukuk) tidak dicentang.")
            edp = st.data_editor(pf[["Simpan", "Tanggal", "Nama", "Rating", "Outlook", "Tingkat", "Teks"]],
                                 hide_index=True, key="ed_pf",
                                 column_config={"Simpan": kol("CheckboxColumn", width="small"),
                                                "Tanggal": kol("DateColumn", format="DD/MM/YYYY"),
                                                "Nama": kol("SelectboxColumn", options=sorted(ent["Entitas"].unique()),
                                                            width="large"),
                                                "Teks": kol("TextColumn", width="large", disabled=True)})
            if st.button("➕ Tambahkan ke tabel Fitch"):
                sel = edp[edp["Simpan"].fillna(False) & edp["Nama"].notna() & edp["Tanggal"].notna()]
                tambah = pd.DataFrame({"Entitas": sel["Nama"], "Tanggal": sel["Tanggal"], "Rating": sel["Rating"],
                                       "Outlook": sel["Outlook"], "Keterangan": sel["Teks"]}, columns=KOLOM_FITCH)
                st.session_state.parsed_fitch = None
                simpan_set(fit=pd.concat([fit, tambah], ignore_index=True), pesan=f"{len(tambah)} rating Fitch ditambahkan.")
                st.rerun()
    ed_fit = st.data_editor(fit, num_rows="dynamic", hide_index=True, key="ed_fit",
                            height=min(40 + 35 * (len(fit) + 2), 420),
                            column_config={"Entitas": kol("SelectboxColumn", options=sorted(ent["Entitas"].unique()),
                                                          width="large"),
                                           "Tanggal": kol("DateColumn", format="DD/MM/YYYY"),
                                           "Rating": kol("TextColumn", width="small", help="Contoh: AAA(idn), AA-(idn)"),
                                           "Keterangan": kol("TextColumn", width="large")})
    if st.button("💾 Simpan tabel Fitch"):
        simpan_set(fit=ed_fit, pesan="Tabel Fitch tersimpan.")
        st.rerun()

# ============================================================================
# TAB 4 – PANDUAN
# ============================================================================
with tab_panduan:
    st.markdown("""
**Cara pakai**
1. Pilih **tanggal periode**.
2. (Opsional) unggah file Excel limit terbaru. Kalau tidak, dipakai file terakhir yang pernah diunggah.
3. Klik **Cek rating**. Untuk tiap counterparty:
   - entitas utama → **PEFINDO** (rating terakhir yang terbit ≤ tanggal periode);
   - kalau tidak ada → **Fitch** (tabel Rating Fitch);
   - kalau dua-duanya tidak ada → **perusahaan anggota group** berikutnya (mis. Salim Group → Indofood → Salim Ivomas);
   - kalau tetap tidak ada → rating di file tidak diubah.
4. Periksa hasilnya (kolom *Rating periode ini* bisa diedit), lalu **Download Excel**.

**Cari Rating** (tab kedua): untuk perusahaan di luar daftar default. Ketik nama/kode → aplikasi memindai
daftar emiten PEFINDO → pilih → tampil rating pada tanggal yang dipilih, rating terkini, riwayat, dan link ke
halaman PEFINDO. Hasilnya bisa langsung ditambahkan ke daftar entitas sebuah group. Untuk Fitch tersedia link
pencarian ke situs Fitch dan form untuk mencatat ratingnya.

**Contoh tanggal**: PEFINDO menerbitkan rating baru 9/11/2026. Jika Anda memilih 8/11/2026, yang diambil adalah
rating sebelum 9/11/2026.

**Catatan**
- Rating PEFINDO diambil dari halaman publik pefindo.com dan disimpan di `data/cache_pefindo/` (dipakai ulang
  di hari yang sama, dan sebagai cadangan bila internet terputus).
- Rating yang sudah ditarik, atau yang terakhir diperbarui lebih dari ±13 bulan sebelum tanggal periode, dianggap
  tidak berlaku → aplikasi mencari ke sumber berikutnya.
- Pengaturan entitas & rating Fitch tersimpan di `data/pengaturan.xlsx` (bisa juga diedit langsung di Excel).
""")

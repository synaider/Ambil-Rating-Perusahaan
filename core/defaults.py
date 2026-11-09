"""
Daftar entitas per group (urutan = prioritas pengambilan rating).

Untuk tiap group, aplikasi mencoba entitas urutan 1: PEFINDO dulu, kalau tidak ada → Fitch.
Kalau dua-duanya tidak ada, lanjut ke entitas urutan berikutnya (perusahaan anggota group), dst.

Kode PEFINDO diambil dari daftar perusahaan di
https://www.pefindo.com/rating-action-reports/rating-history (dicek 9 Nov 2026).
Bisa diubah di aplikasi (tab "Sumber Rating") atau langsung di data/pengaturan.xlsx.
"""
from __future__ import annotations

import re

# (Group, [(Entitas, Kode PEFINDO, Alias untuk mencocokkan teks rilis Fitch[, Lembaga]), ...])
# Lembaga: "Otomatis" (PEFINDO dulu, lalu Fitch) | "PEFINDO" (hanya PEFINDO) | "Fitch" (hanya Fitch)
ENTITAS_GROUP: list[tuple[str, list[tuple]]] = [
    ("BRI GROUP", [("PT Bank Rakyat Indonesia (Persero) Tbk", "BBRI", "Bank Rakyat Indonesia; BBRI"),
                   ("PT Bank Raya Indonesia Tbk", "AGRO", "Bank Raya Indonesia"),
                   ("PT BRI Multifinance Indonesia", "BRIF", "BRI Multifinance"),
                   ("PT Pegadaian", "PPGD", "Pegadaian"),
                   ("PT Permodalan Nasional Madani", "PNMP", "Permodalan Nasional Madani")]),
    ("MANDIRI GROUP", [("PT Bank Mandiri (Persero) Tbk", "BMRI", "Bank Mandiri; BMRI"),
                       ("PT Bank Mandiri Taspen", "BMTP", "Bank Mandiri Taspen"),
                       ("PT Mandiri Tunas Finance", "TUFI", "Mandiri Tunas Finance"),
                       ("PT Mandiri Utama Finance", "MUFI", "Mandiri Utama Finance")]),
    ("PLN GROUP", [("PT Perusahaan Listrik Negara (Persero)", "PPLN", "Perusahaan Listrik Negara; PLN"),
                   ("PT Indonesia Power", "IDPW", "Indonesia Power"),
                   ("PT PLN Nusantara Renewables", "PLNR", "PLN Nusantara Renewables")]),
    ("SMI GROUP", [("PT Sarana Multi Infrastruktur (Persero)", "SMII", "Sarana Multi Infrastruktur")]),
    ("INDOSAT GROUP", [("PT Indosat Tbk", "ISAT", "Indosat")]),
    ("BNI GROUP", [("PT Bank Negara Indonesia (Persero) Tbk", "BBNI", "Bank Negara Indonesia; BBNI")]),
    ("DJARUM GROUP", [("PT Bank Central Asia Tbk", "BBCA", "Bank Central Asia; BBCA"),
                      ("PT BCA Finance", "BCAF", "BCA Finance"),
                      ("PT BCA Multi Finance", "BCAM", "BCA Multi Finance"),
                      ("PT Bank BCA Syariah", "BCAS", "BCA Syariah"),
                      ("PT Bank Digital BCA", "BCAD", "Bank Digital BCA")]),
    ("INDONESIA EXIMBANK GROUP", [("Lembaga Pembiayaan Ekspor Indonesia", "BEXI",
                                   "Lembaga Pembiayaan Ekspor Indonesia; Indonesia Eximbank; LPEI")]),
    ("SMF GROUP", [("PT Sarana Multigriya Finansial (Persero)", "SMFP", "Sarana Multigriya Finansial")]),
    ("MAYBANK INDONESIA GROUP", [("PT Bank Maybank Indonesia Tbk", "BNII", "Maybank Indonesia; BNII"),
                                 ("PT Maybank Indonesia Finance", "BIIF", "Maybank Indonesia Finance")]),
    ("ASTRA GROUP", [("PT Astra International Tbk", "ASII", "Astra International"),
                     ("PT Astra Sedaya Finance", "ASDF", "Astra Sedaya Finance"),
                     ("PT Federal International Finance", "FIFA", "Federal International Finance"),
                     ("PT Toyota Astra Financial Services", "TAFS", "Toyota Astra Financial Services"),
                     ("PT Astra Auto Finance", "AAFN", "Astra Auto Finance"),
                     ("PT Komatsu Astra Finance", "KAFN", "Komatsu Astra Finance")]),
    ("TELKOM GROUP", [("PT Telkom Indonesia (Persero) Tbk", "TLKM", "Telkom Indonesia; Telekomunikasi Indonesia"),
                      ("PT Dayamitra Telekomunikasi Tbk", "MTEL", "Dayamitra Telekomunikasi; Mitratel")]),
    ("CIMB NIAGA GROUP", [("PT Bank CIMB Niaga Tbk", "BNGA", "CIMB Niaga")]),
    ("MITSUBISHI UFJ F.GROUP", [("PT Bank Danamon Indonesia Tbk", "BDMN", "Bank Danamon"),
                                ("PT Adira Dinamika Multi Finance Tbk", "ADMF", "Adira Dinamika Multi Finance; Adira Finance")]),
    ("SUMITOMO MITSUI FINANCIAL GROUP", [("PT Bank SMBC Indonesia Tbk", "BTPN", "Bank SMBC Indonesia; SMBC Indonesia; BTPN"),
                                         ("PT Oto Multiartha", "OTMA", "Oto Multiartha"),
                                         ("PT Summit Oto Finance", "SOFN", "Summit Oto Finance")]),
    ("PT WOORI FINANCIAL GROUP", [("PT Bank Woori Saudara Indonesia 1906 Tbk", "SDRA", "Woori Saudara; Bank Woori")]),
    ("BTN GROUP", [("PT Bank Tabungan Negara (Persero) Tbk", "BBTN", "Bank Tabungan Negara; BBTN")]),
    ("SEMEN INDONESIA GROUP", [("PT Semen Indonesia (Persero) Tbk", "SMGR", "Semen Indonesia"),
                               ("PT Semen Baturaja Tbk", "SMBR", "Semen Baturaja")]),
    ("INJOURNEY GROUP", [("PT Aviasi Pariwisata Indonesia (Persero)", "INJH", "Aviasi Pariwisata Indonesia; InJourney"),
                         ("PT Angkasa Pura Indonesia", "APIA", "Angkasa Pura Indonesia")]),
    ("KAI GROUP", [("PT Kereta Api Indonesia (Persero)", "KAII", "Kereta Api Indonesia"),
                   ("PT Kereta Commuter Indonesia", "KCIN", "Kereta Commuter Indonesia"),
                   ("PT Kereta Api Logistik", "KALO", "Kereta Api Logistik")]),
    ("BANGKOK BANK PC LIMITED", [("Bangkok Bank Public Company Limited", "", "Bangkok Bank")]),
    ("SALIM GROUP", [("PT Indofood Sukses Makmur Tbk", "INDF", "Indofood Sukses Makmur; Indofood"),
                     ("PT Salim Ivomas Pratama Tbk", "SIMP", "Salim Ivomas Pratama")]),
    ("PANIN GROUP", [("PT Bank Pan Indonesia Tbk", "PNBN", "Bank Pan Indonesia; Panin Bank"),
                     ("PT Clipan Finance Indonesia Tbk", "CFIN", "Clipan Finance"),
                     ("PT Bank Panin Dubai Syariah Tbk", "PNBS", "Panin Dubai Syariah")]),
    ("CT CORP GROUP", [("PT Bank Mega Tbk", "MEGA", "Bank Mega"),
                       ("PT Allo Bank Indonesia Tbk", "BBHI", "Allo Bank"),
                       ("PT Mega Auto Finance", "MAFN", "Mega Auto Finance"),
                       ("PT Mega Central Finance", "MCFN", "Mega Central Finance")]),
    ("HK GROUP", [("PT Hutama Karya (Persero)", "PTHK", "Hutama Karya (Persero)"),
                  ("PT Hutama Karya Infrastruktur", "HKIF", "Hutama Karya Infrastruktur")]),
    ("SINARMAS GROUP", [("PT Asuransi Sinar Mas", "ASMS", "Asuransi Sinar Mas")]),
    ("JASA MARGA GROUP", [("PT Jasa Marga (Persero) Tbk", "JSMR", "Jasa Marga"),
                          ("PT Jasamarga Pandaan Tol", "JPTR", "Jasamarga Pandaan Tol"),
                          ("PT Marga Lingkar Jakarta", "MLJK", "Marga Lingkar Jakarta")]),
    ("BANK RAYA INDONESIA", [("PT Bank Raya Indonesia Tbk", "AGRO", "Bank Raya Indonesia")]),
    ("BANK SULSELBAR GROUP", [("PT Bank Pembangunan Daerah Sulawesi Selatan dan Sulawesi Barat", "BSSB",
                               "Sulawesi Selatan dan Sulawesi Barat; Bank Sulselbar")]),
    ("BARITO PACIFIC GROUP", [("PT Barito Pacific Tbk", "BRPT", "Barito Pacific")]),
    ("BANK JATENG GROUP", [("PT Bank Pembangunan Daerah Jawa Tengah", "BJTG", "Bank Jateng; Jawa Tengah")]),
    ("BANK JATIM GROUP", [("PT Bank Pembangunan Daerah Jawa Timur Tbk", "BJTM", "Bank Jatim; Jawa Timur")]),
    ("BANK NAGARI GROUP", [("PT Bank Nagari", "BSBR", "Bank Nagari")]),
    ("PTPP GROUP", [("PT Pembangunan Perumahan (Persero) Tbk", "PTPP", "Pembangunan Perumahan; PTPP")]),
    ("SULUTGO GROUP", [("PT Bank SulutGo", "BSLT", "SulutGo; Bank Sulut")]),
    ("WASKITA GROUP", [("PT Waskita Karya (Persero) Tbk", "WSKT", "Waskita Karya"),
                       ("PT Waskita Beton Precast Tbk", "WSBP", "Waskita Beton Precast"),
                       ("PT Waskita Toll Road", "WSTR", "Waskita Toll Road")]),
    ("DBS GROUP", [("PT Bank DBS Indonesia", "", "Bank DBS Indonesia; DBS Indonesia")]),
    ("BAF GROUP", [("PT Bussan Auto Finance", "BAFI", "Bussan Auto Finance")]),
    ("BJB GROUP", [("PT Bank Pembangunan Daerah Jawa Barat dan Banten Tbk", "BJBR", "Jawa Barat dan Banten; Bank BJB")]),
    ("MERDEKA GROUP", [("PT Merdeka Copper Gold Tbk", "MDKA", "Merdeka Copper Gold"),
                       ("PT Merdeka Battery Materials Tbk", "MBMA", "Merdeka Battery Materials")]),
    ("OCBC GROUP", [("PT Bank OCBC NISP Tbk", "NISP", "OCBC NISP; OCBC Indonesia")]),
    ("PT METAMORFOSIS EKOSISTEM INDONESIA", [("PT Metamorfosis Ekosistem Indonesia", "", "Metamorfosis Ekosistem")]),
    ("BANK JAKARTA GROUP", [("PT Bank Jakarta", "BDKI", "Bank Jakarta; Bank DKI")]),
    ("BANK BTPN GROUP", [("PT Bank SMBC Indonesia Tbk", "BTPN", "Bank SMBC Indonesia; SMBC Indonesia; BTPN", "Fitch")]),
]


# Group yang sumbernya ditetapkan pengguna (9 Nov 2026). Pengaturan lama otomatis disesuaikan sekali.
SUMBER_DITETAPKAN = {"SINARMAS GROUP", "BARITO PACIFIC GROUP", "BANK BTPN GROUP"}


def key_group(nama: str) -> str:
    """'BRI GROUP+)' -> 'BRI GROUP'; 'TELKOM  GROUP' -> 'TELKOM GROUP'; 'BAF Group' -> 'BAF GROUP'."""
    s = re.sub(r"[+\-]\)\s*$", "", str(nama or "")).strip()
    return re.sub(r"\s+", " ", s).upper()


def default_untuk(nama: str) -> tuple[str, str, str, str]:
    """Kompatibilitas modul lama: (kode, entitas, alias, catatan) entitas urutan pertama."""
    k = key_group(nama)
    for g, ents in ENTITAS_GROUP:
        if g == k and ents:
            e, kode, alias = ents[0][:3]
            return kode, e, alias, ""
    return "", "", "", ""

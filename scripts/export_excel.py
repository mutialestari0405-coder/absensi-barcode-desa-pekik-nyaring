#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Export Excel berwarna Rekap Absensi Perangkat Desa Pekik Nyaring.
Gaya seperti template attendance list: grid bulanan berwarna, ringkasan
berwarna, dan grafik Excel native (bar / line / pie).

Pemakaian: python3 export_excel.py payload.json output.xlsx
Payload: {"dari": "YYYY-MM-DD", "sampai": "YYYY-MM-DD",
          "perangkat": [...], "absensi": [...]}
"""
import calendar
import json
import sys
from datetime import date

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------- Palet warna (senada branding hijau-emas desa) ----------
HIJAU_TUA = "17492C"
HIJAU = "1E6339"
HIJAU_MED = "267D47"
HIJAU_MUDA = "DCEFE2"
HIJAU_SOFT = "F0F8F3"
EMAS = "C9A227"
EMAS_MUDA = "F7EFD4"
KREM = "FAF9F4"
MERAH = "C0392B"
MERAH_MUDA = "FDECEA"
ABU = "E8E8E8"
ABU_MUDA = "F5F5F5"
ORANYE_MUDA = "FDEBD0"
BIRU = "1565C0"
BIRU_MUDA = "DBEAFE"

# Kode keterangan kehadiran (selain Masuk) untuk grid bulanan
KODE_KETERANGAN = {"IZIN": "I", "SAKIT": "S", "DINAS_LUAR": "DL", "ALPA": "A"}
LABEL_KETERANGAN = {"MASUK": "Masuk", "IZIN": "Izin", "SAKIT": "Sakit",
                    "DINAS_LUAR": "Dinas Luar", "ALPA": "Alpa"}

FILL = lambda c: PatternFill("solid", fgColor=c)
GARIS = Side(style="thin", color="B9C9BE")
BORDER = Border(left=GARIS, right=GARIS, top=GARIS, bottom=GARIS)
FONT_BASE = "Calibri"

HARI_ID = ["Sen", "Sel", "Rab", "Kam", "Jum", "Sab", "Min"]  # weekday() 0=Senin
BULAN_ID = ["JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", "JULI",
            "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"]


def fjudul(ws, teks, subteks, kolom_akhir):
    """Banner judul hijau-emas seperti header template."""
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=kolom_akhir)
    c = ws.cell(row=1, column=1, value=teks)
    c.font = Font(name=FONT_BASE, bold=True, size=16, color=EMAS)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for col in range(1, kolom_akhir + 1):
        ws.cell(row=1, column=col).fill = FILL(HIJAU_TUA)
    ws.row_dimensions[1].height = 34

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=kolom_akhir)
    c2 = ws.cell(row=2, column=2 - 1, value=subteks)
    c2.font = Font(name=FONT_BASE, size=10, color="FFFFFF")
    c2.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for col in range(1, kolom_akhir + 1):
        ws.cell(row=2, column=col).fill = FILL(HIJAU)
    ws.row_dimensions[2].height = 18


def legenda(ws, baris, kolom_awal):
    """Baris legenda kecil di bawah judul."""
    item = [
        ("DP = Datang & Pulang", FILL(HIJAU_MUDA), HIJAU_TUA),
        ("D = Hanya Datang", FILL(EMAS_MUDA), "7A6210"),
        ("I = Izin • S = Sakit", FILL(ORANYE_MUDA), "8A5A00"),
        ("DL = Dinas Luar • A = Alpa", FILL(BIRU_MUDA), BIRU),
        ("(kosong) = Tidak absen", FILL(MERAH_MUDA), MERAH),
        ("Kolom abu-abu = akhir pekan", FILL(ABU), "555555"),
    ]
    col = kolom_awal
    for teks, fill, warna in item:
        c = ws.cell(row=baris, column=col, value=teks)
        c.font = Font(name=FONT_BASE, size=8, color=warna, italic=True)
        c.fill = fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells(start_row=baris, start_column=col, end_row=baris, end_column=col + 2)
        for k in range(col, col + 3):
            ws.cell(row=baris, column=k).fill = fill
        col += 3


def sheet_rekap_bulanan(wb, perangkat, absensi, tahun, bulan, hari_ini):
    ws = wb.active
    ws.title = "Rekap Bulanan"
    ws.sheet_view.showGridLines = False

    jml_hari = calendar.monthrange(tahun, bulan)[1]
    hari_efektif_max = hari_ini if (tahun, bulan) == (hari_ini.year, hari_ini.month) else date(tahun, bulan, jml_hari)

    # Peta absensi: (perangkatId, tanggal) -> 'DP' | 'D' | 'I' | 'S' | 'DL' | 'A'
    peta = {}
    for a in absensi:
        tgl = a["tanggal"]  # YYYY-MM-DD
        ket = (a.get("keteranganKehadiran") or "MASUK").upper()
        if ket in KODE_KETERANGAN:
            peta[(a["perangkatId"], tgl)] = KODE_KETERANGAN[ket]
        else:
            peta[(a["perangkatId"], tgl)] = "DP" if a.get("jamPulang") else "D"

    kolom_awal_hari = 4  # A=No, B=Nama, C=Jabatan
    kolom_akhir = kolom_awal_hari + jml_hari + 3  # + Hadir/Lengkap/Tidak

    fjudul(ws, f"REKAP ABSENSI PERANGKAT DESA — {BULAN_ID[bulan - 1]}-{tahun}",
           "Pemerintah Desa Pekik Nyaring  •  Kec. Pondok Kelapa, Kab. Bengkulu Tengah", kolom_akhir)
    legenda(ws, 3, 1)

    # ---------- Header tabel ----------
    r_no, r_hari = 4, 5
    header_tengah = Alignment(horizontal="center", vertical="center")
    for col, teks in ((1, "No"), (2, "Nama Perangkat Desa"), (3, "Jabatan")):
        ws.merge_cells(start_row=r_no, start_column=col, end_row=r_hari, end_column=col)
        c = ws.cell(row=r_no, column=col, value=teks)
        c.font = Font(name=FONT_BASE, bold=True, size=10, color="FFFFFF")
        c.fill = FILL(HIJAU_TUA)
        c.alignment = header_tengah
        for r in (r_no, r_hari):
            ws.cell(row=r, column=col).fill = FILL(HIJAU_TUA)
            ws.cell(row=r, column=col).border = BORDER
    for d in range(1, jml_hari + 1):
        col = kolom_awal_hari + d - 1
        wk = date(tahun, bulan, d).weekday()
        weekend = wk >= 5
        c1 = ws.cell(row=r_no, column=col, value=d)
        c2 = ws.cell(row=r_hari, column=col, value=HARI_ID[wk])
        for c in (c1, c2):
            c.font = Font(name=FONT_BASE, bold=True, size=9,
                          color="8A8A8A" if weekend else "FFFFFF")
            c.fill = FILL(ABU if weekend else HIJAU_TUA)
            c.alignment = header_tengah
            c.border = BORDER
    for offset, (teks, warna_fill, warna_font) in enumerate([
        ("Hadir", HIJAU_MED, "FFFFFF"),
        ("Lengkap D&P", HIJAU_TUA, "FFFFFF"),
        ("Tidak Hadir", MERAH, "FFFFFF"),
    ]):
        col = kolom_awal_hari + jml_hari + offset
        ws.merge_cells(start_row=r_no, start_column=col, end_row=r_hari, end_column=col)
        c = ws.cell(row=r_no, column=col, value=teks)
        c.font = Font(name=FONT_BASE, bold=True, size=9, color=warna_font)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for r in (r_no, r_hari):
            ws.cell(row=r, column=col).fill = FILL(warna_fill)
            ws.cell(row=r, column=col).border = BORDER
    ws.row_dimensions[r_no].height = 20
    ws.row_dimensions[r_hari].height = 16

    # ---------- Baris per perangkat ----------
    ringkas = []  # (nama, jabatan, hadir, lengkap, tidak)
    r = r_hari + 1
    for idx, p in enumerate(perangkat, start=1):
        zebra = KREM if idx % 2 == 0 else "FFFFFF"
        c = ws.cell(row=r, column=1, value=idx)
        c.font = Font(name=FONT_BASE, size=9, color="666666")
        c.alignment = header_tengah
        c.fill = FILL(zebra)
        c.border = BORDER
        c = ws.cell(row=r, column=2, value=p["nama"])
        c.font = Font(name=FONT_BASE, bold=True, size=9, color=HIJAU_TUA)
        c.alignment = Alignment(vertical="center")
        c.fill = FILL(zebra)
        c.border = BORDER
        c = ws.cell(row=r, column=3, value=p["jabatan"])
        c.font = Font(name=FONT_BASE, size=8, color="555555")
        c.alignment = Alignment(vertical="center")
        c.fill = FILL(zebra)
        c.border = BORDER

        hadir = lengkap = tidak = 0
        for d in range(1, jml_hari + 1):
            col = kolom_awal_hari + d - 1
            tgl = f"{tahun:04d}-{bulan:02d}-{d:02d}"
            wk = date(tahun, bulan, d).weekday()
            weekend = wk >= 5
            kode = peta.get((p["id"], tgl))
            cell = ws.cell(row=r, column=col)
            if kode:
                if kode == "DP":
                    cell.value = "DP"
                    cell.fill = FILL(HIJAU_MUDA)
                    cell.font = Font(name=FONT_BASE, bold=True, size=8, color=HIJAU_TUA)
                    lengkap += 1
                    hadir += 1
                elif kode == "D":
                    cell.value = "D"
                    cell.fill = FILL(EMAS_MUDA)
                    cell.font = Font(name=FONT_BASE, bold=True, size=8, color="7A6210")
                    hadir += 1
                elif kode == "I":
                    cell.value = "I"
                    cell.fill = FILL(ORANYE_MUDA)
                    cell.font = Font(name=FONT_BASE, bold=True, size=8, color="8A5A00")
                elif kode == "S":
                    cell.value = "S"
                    cell.fill = FILL(MERAH_MUDA)
                    cell.font = Font(name=FONT_BASE, bold=True, size=8, color=MERAH)
                elif kode == "DL":
                    cell.value = "DL"
                    cell.fill = FILL(BIRU_MUDA)
                    cell.font = Font(name=FONT_BASE, bold=True, size=8, color=BIRU)
                else:  # A = Alpa
                    cell.value = "A"
                    cell.fill = FILL(MERAH)
                    cell.font = Font(name=FONT_BASE, bold=True, size=8, color="FFFFFF")
            elif not weekend and date(tahun, bulan, d) <= hari_efektif_max:
                tidak += 1
                cell.fill = FILL(MERAH_MUDA)
            elif weekend:
                cell.fill = FILL(ABU_MUDA)
            cell.alignment = header_tengah
            cell.border = BORDER

        for offset, (nilai, fill, warna) in enumerate([
            (hadir, HIJAU_MUDA, HIJAU_TUA),
            (lengkap, HIJAU_TUA, "FFFFFF"),
            (tidak, MERAH_MUDA if tidak else zebra, MERAH if tidak else "999999"),
        ]):
            col = kolom_awal_hari + jml_hari + offset
            c = ws.cell(row=r, column=col, value=nilai)
            c.font = Font(name=FONT_BASE, bold=True, size=10, color=warna)
            c.fill = FILL(fill)
            c.alignment = header_tengah
            c.border = BORDER
        ringkas.append((p["nama"], p["jabatan"], hadir, lengkap, tidak))
        ws.row_dimensions[r].height = 18
        r += 1

    # ---------- Band ringkasan per hari (seperti tabel bawah template) ----------
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=kolom_akhir)
    c = ws.cell(row=r, column=1, value="REKAP KEHADIRAN PER HARI")
    c.font = Font(name=FONT_BASE, bold=True, size=10, color="FFFFFF")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for col in range(1, kolom_akhir + 1):
        ws.cell(row=r, column=col).fill = FILL(HIJAU)
    r += 1
    r_tgl = r
    c = ws.cell(row=r, column=1, value="Tanggal")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    c.font = Font(name=FONT_BASE, bold=True, size=9, color="FFFFFF")
    c.fill = FILL(HIJAU_TUA)
    c.alignment = header_tengah
    for cc in range(1, 4):
        ws.cell(row=r, column=cc).fill = FILL(HIJAU_TUA)
        ws.cell(row=r, column=cc).border = BORDER
    for d in range(1, jml_hari + 1):
        col = 3 + d
        wk = date(tahun, bulan, d).weekday()
        weekend = wk >= 5
        c = ws.cell(row=r, column=col, value=d)
        c.font = Font(name=FONT_BASE, bold=True, size=8, color="8A8A8A" if weekend else "FFFFFF")
        c.fill = FILL(ABU if weekend else HIJAU_TUA)
        c.alignment = header_tengah
        c.border = BORDER
    r += 1
    c = ws.cell(row=r, column=1, value="Jumlah Hadir")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    c.font = Font(name=FONT_BASE, bold=True, size=9, color=HIJAU_TUA)
    c.fill = FILL(EMAS_MUDA)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for cc in range(1, 4):
        ws.cell(row=r, column=cc).fill = FILL(EMAS_MUDA)
        ws.cell(row=r, column=cc).border = BORDER
    for d in range(1, jml_hari + 1):
        col = 3 + d
        tgl = f"{tahun:04d}-{bulan:02d}-{d:02d}"
        n = sum(1 for a in absensi if a["tanggal"] == tgl)
        c = ws.cell(row=r, column=col, value=n)
        c.font = Font(name=FONT_BASE, bold=True, size=8,
                      color=HIJAU_TUA if n else MERAH)
        c.fill = FILL(HIJAU_MUDA if n else MERAH_MUDA)
        c.alignment = header_tengah
        c.border = BORDER

    # ---------- Lebar kolom & freeze ----------
    ws.column_dimensions["A"].width = 4.5
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 20
    for d in range(1, jml_hari + 1):
        ws.column_dimensions[get_column_letter(3 + d)].width = 4.6
    for offset in range(3):
        ws.column_dimensions[get_column_letter(3 + jml_hari + 1 + offset)].width = 8.5
    ws.freeze_panes = ws.cell(row=r_hari + 1, column=4)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    return ringkas


def sheet_grafik(wb, ringkas, absensi, tahun, bulan):
    ws = wb.create_sheet("Grafik")
    ws.sheet_view.showGridLines = False
    jml_hari = calendar.monthrange(tahun, bulan)[1]

    ws.merge_cells("A1:L1")
    c = ws.cell(row=1, column=1, value=f"GRAFIK ABSENSI — {BULAN_ID[bulan - 1]} {tahun}")
    c.font = Font(name=FONT_BASE, bold=True, size=13, color="FFFFFF")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for col in range(1, 13):
        ws.cell(row=1, column=col).fill = FILL(HIJAU_TUA)
    ws.row_dimensions[1].height = 26

    # Data 1: per perangkat
    c = ws.cell(row=3, column=1, value="Kehadiran per Perangkat")
    c.font = Font(name=FONT_BASE, bold=True, size=10, color=HIJAU_TUA)
    ws.cell(row=4, column=1, value="Nama").font = Font(name=FONT_BASE, bold=True, size=9)
    ws.cell(row=4, column=2, value="Hadir").font = Font(name=FONT_BASE, bold=True, size=9)
    ws.cell(row=4, column=3, value="Lengkap D&P").font = Font(name=FONT_BASE, bold=True, size=9)
    for i, (nama, _, hadir, lengkap, _) in enumerate(ringkas):
        ws.cell(row=5 + i, column=1, value=nama).font = Font(name=FONT_BASE, size=9)
        ws.cell(row=5 + i, column=2, value=hadir).font = Font(name=FONT_BASE, size=9)
        ws.cell(row=5 + i, column=3, value=lengkap).font = Font(name=FONT_BASE, size=9)

    bar = BarChart()
    bar.type = "col"
    bar.grouping = "clustered"
    bar.title = "Total Kehadiran per Perangkat Desa"
    bar.y_axis.title = "Jumlah Hari"
    bar.x_axis.title = None
    bar.height = 9.5
    bar.width = 24
    data = Reference(ws, min_col=2, min_row=4, max_col=3, max_row=4 + len(ringkas))
    kats = Reference(ws, min_col=1, min_row=5, max_row=4 + len(ringkas))
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(kats)
    bar.series[0].graphicalProperties.solidFill = HIJAU_MED
    bar.series[1].graphicalProperties.solidFill = EMAS
    bar.gapWidth = 60
    ws.add_chart(bar, "E4")

    # Data 2: per hari
    baris2 = 26
    c = ws.cell(row=baris2, column=1, value="Kehadiran per Hari")
    c.font = Font(name=FONT_BASE, bold=True, size=10, color=HIJAU_TUA)
    ws.cell(row=baris2 + 1, column=1, value="Tanggal").font = Font(name=FONT_BASE, bold=True, size=9)
    ws.cell(row=baris2 + 1, column=2, value="Jumlah Hadir").font = Font(name=FONT_BASE, bold=True, size=9)
    for d in range(1, jml_hari + 1):
        tgl = f"{tahun:04d}-{bulan:02d}-{d:02d}"
        n = sum(1 for a in absensi if a["tanggal"] == tgl)
        ws.cell(row=baris2 + 1 + d, column=1, value=d).font = Font(name=FONT_BASE, size=9)
        ws.cell(row=baris2 + 1 + d, column=2, value=n).font = Font(name=FONT_BASE, size=9)

    garis = LineChart()
    garis.title = "Jumlah Kehadiran per Hari"
    garis.y_axis.title = "Orang"
    garis.x_axis.title = "Tanggal"
    garis.height = 9.5
    garis.width = 24
    data = Reference(ws, min_col=2, min_row=baris2 + 1, max_row=baris2 + 1 + jml_hari)
    kats = Reference(ws, min_col=1, min_row=baris2 + 2, max_row=baris2 + 1 + jml_hari)
    garis.add_data(data, titles_from_data=True)
    garis.set_categories(kats)
    garis.series[0].graphicalProperties.line.solidFill = HIJAU
    garis.series[0].graphicalProperties.line.width = 28000
    garis.series[0].smooth = False
    ws.add_chart(garis, f"E{baris2}")

    # Data 3: komposisi (pie)
    baris3 = baris2 + jml_hari + 5
    c = ws.cell(row=baris3, column=1, value="Komposisi Absensi (perangkat × hari efektif)")
    c.font = Font(name=FONT_BASE, bold=True, size=10, color=HIJAU_TUA)
    total_dp = sum(1 for a in absensi if a.get("jamPulang"))
    total_d = sum(1 for a in absensi if not a.get("jamPulang"))
    total_efektif = sum(h[4] for h in ringkas) + total_dp + total_d
    label = [("Lengkap Datang & Pulang", total_dp), ("Hanya Datang", total_d), ("Tidak Hadir", sum(h[4] for h in ringkas))]
    ws.cell(row=baris3 + 1, column=1, value="Kategori").font = Font(name=FONT_BASE, bold=True, size=9)
    ws.cell(row=baris3 + 1, column=2, value="Jumlah").font = Font(name=FONT_BASE, bold=True, size=9)
    for i, (k, v) in enumerate(label):
        ws.cell(row=baris3 + 2 + i, column=1, value=k).font = Font(name=FONT_BASE, size=9)
        ws.cell(row=baris3 + 2 + i, column=2, value=v).font = Font(name=FONT_BASE, size=9)

    pie = PieChart()
    pie.title = "Komposisi Kehadiran"
    pie.height = 9.5
    pie.width = 12
    data = Reference(ws, min_col=2, min_row=baris3 + 1, max_row=baris3 + 4)
    kats = Reference(ws, min_col=1, min_row=baris3 + 2, max_row=baris3 + 4)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(kats)
    from openpyxl.chart.series import DataPoint
    slice_warna = [HIJAU_MED, EMAS, MERAH]
    for i, wc in enumerate(slice_warna):
        dp = DataPoint(idx=i)
        dp.graphicalProperties.solidFill = wc
        pie.series[0].data_points.append(dp)
    ws.add_chart(pie, f"E{baris3}")

    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 14
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def gaya_tabel_detail(ws, header, rows, lebar, hyperlink_col=None):
    """Tabel detail ber-header hijau + zebra striping + border."""
    ws.sheet_view.showGridLines = False
    for i, h in enumerate(header, start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = Font(name=FONT_BASE, bold=True, size=9, color="FFFFFF")
        c.fill = FILL(HIJAU_TUA)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[1].height = 24
    for r, baris in enumerate(rows, start=2):
        zebra = HIJAU_SOFT if r % 2 == 0 else "FFFFFF"
        for i, nilai in enumerate(baris, start=1):
            c = ws.cell(row=r, column=i, value=nilai)
            c.font = Font(name=FONT_BASE, size=9)
            c.fill = FILL(zebra)
            c.border = BORDER
            c.alignment = Alignment(vertical="center", wrap_text=False)
    for i, w in enumerate(lebar, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    if hyperlink_col:
        for r in range(2, len(rows) + 2):
            c = ws.cell(row=r, column=hyperlink_col)
            if c.value and str(c.value).startswith("http"):
                c.hyperlink = c.value
                c.value = "Buka Peta"
                c.font = Font(name=FONT_BASE, size=9, color="1565C0", underline="single")


def sheet_riwayat(wb, absensi):
    ws = wb.create_sheet("Riwayat Absensi")
    header = ["No", "Tanggal", "NIPD", "Nama", "Jabatan", "Status", "Jam Datang",
              "Jam Pulang", "Kategori", "Keterangan", "Metode", "GPS Datang", "GPS Pulang"]
    rows = []
    for i, a in enumerate(absensi, start=1):
        gps_d = (f"https://www.google.com/maps?q={a['latitude']},{a['longitude']}"
                 if a.get("latitude") is not None and a.get("longitude") is not None else None)
        gps_p = (f"https://www.google.com/maps?q={a['latPulang']},{a['lonPulang']}"
                 if a.get("latPulang") is not None and a.get("lonPulang") is not None else None)
        ket = (a.get("keteranganKehadiran") or "MASUK").upper()
        rows.append([
            i, a["tanggal"], a.get("nipd", "-"), a.get("nama", "-"), a.get("jabatan", "-"),
            a.get("status", "-"), a.get("jamDatang") or "-", a.get("jamPulang") or "-",
            "AKTIF" if a["kategori"] == "AKTIF" else "PASIF",
            LABEL_KETERANGAN.get(ket, ket), a.get("metode", "-"),
            gps_d or "Tidak ada GPS", gps_p or "-",
        ])
    gaya_tabel_detail(ws, header, rows,
                      [4.5, 11, 9, 22, 24, 8, 10, 10, 9, 11.5, 8.5, 11, 11],
                      hyperlink_col=12)
    # Beri warna kolom kategori & keterangan
    gaya_ket = {
        "Masuk": (HIJAU_MUDA, HIJAU_TUA), "Izin": (ORANYE_MUDA, "8A5A00"),
        "Sakit": (MERAH_MUDA, MERAH), "Dinas Luar": (BIRU_MUDA, BIRU),
        "Alpa": (MERAH, "FFFFFF"),
    }
    for r in range(2, len(rows) + 2):
        c = ws.cell(row=r, column=9)
        if c.value == "AKTIF":
            c.fill = FILL(HIJAU_MUDA)
            c.font = Font(name=FONT_BASE, size=9, bold=True, color=HIJAU_TUA)
        else:
            c.fill = FILL(EMAS_MUDA)
            c.font = Font(name=FONT_BASE, size=9, bold=True, color="7A6210")
        ck = ws.cell(row=r, column=10)
        fill_ket, warna_ket = gaya_ket.get(str(ck.value), (HIJAU_SOFT, "555555"))
        ck.fill = FILL(fill_ket)
        ck.font = Font(name=FONT_BASE, size=9, bold=True, color=warna_ket)
        ck.alignment = Alignment(horizontal="center", vertical="center")


def sheet_perangkat(wb, perangkat):
    ws = wb.create_sheet("Data Perangkat Desa")
    header = ["No", "NIPD", "Nama Lengkap", "NIK", "Jabatan", "Jenis Kelamin",
              "Pendidikan", "Nomor SK", "Status", "Keterangan"]
    rows = []
    for i, p in enumerate(perangkat, start=1):
        rows.append([i, p["nipd"], p["nama"], p.get("nik") or "-", p["jabatan"],
                     p.get("jenisKelamin", "-"), p.get("pendidikan") or "-",
                     p.get("nomorSk") or "-", p["status"], p.get("keterangan") or "-"])
    gaya_tabel_detail(ws, header, rows,
                      [4.5, 9.5, 26, 20, 27, 13, 11, 10, 9, 34])
    for r in range(2, len(rows) + 2):
        c = ws.cell(row=r, column=9)
        if c.value == "AKTIF":
            c.fill = FILL(HIJAU_MUDA)
            c.font = Font(name=FONT_BASE, size=9, bold=True, color=HIJAU_TUA)
        else:
            c.fill = FILL(EMAS_MUDA)
            c.font = Font(name=FONT_BASE, size=9, bold=True, color="7A6210")


def main():
    payload = json.load(open(sys.argv[1], encoding="utf-8"))
    perangkat = payload["perangkat"]
    absensi = payload["absensi"]
    dari = payload.get("dari") or (absensi[0]["tanggal"] if absensi else f"{date.today().year}-{date.today().month:02d}-01")
    tahun, bulan = int(dari[:4]), int(dari[5:7])
    hari_ini = date.today()

    wb = Workbook()
    ringkas = sheet_rekap_bulanan(wb, perangkat, absensi, tahun, bulan, hari_ini)
    sheet_grafik(wb, ringkas, absensi, tahun, bulan)
    sheet_riwayat(wb, absensi)
    sheet_perangkat(wb, perangkat)

    wb.save(sys.argv[2])
    print(f"OK: {sys.argv[2]}")


if __name__ == "__main__":
    main()

from io import BytesIO

from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT

from services.helpers import set_font


def set_cell_text(cell, text, size=8, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    r = p.add_run(str(text))
    set_font(r, size=size, bold=bold)


def set_col_widths(table):
    widths = [
        Cm(1.0),    # No Urut
        Cm(3.2),    # No Surat
        Cm(10.0),   # Ditujukan
        Cm(2.2),    # No X5
        Cm(1.8),    # Berat
        Cm(2.0),    # Tarif
        Cm(3.0),    # Keterangan
    ]

    for row in table.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = width


def build_daftar_pengiriman(data):
    doc = Document()

    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width

    section.top_margin = Cm(1.2)
    section.bottom_margin = Cm(1.2)
    section.left_margin = Cm(1.2)
    section.right_margin = Cm(1.2)

    # Judul
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("DAFTAR PENGIRIMAN POS SURAT DINAS (X5)")
    set_font(r, size=11, bold=True)

    rows_data = data.get("selected_surat", [])

    table = doc.add_table(rows=1, cols=7)
    table.style = "Table Grid"
    table.autofit = False

    headers = [
        "No\nUrut",
        "No. Surat",
        "Ditujukan",
        "No.X5",
        "Berat (Gr)",
        "Tarif (Rp.)",
        "Keterangan",
    ]

    for i, header in enumerate(headers):
        set_cell_text(
            table.rows[0].cells[i],
            header,
            size=8,
            bold=True,
            align=WD_ALIGN_PARAGRAPH.CENTER
        )

    for idx, row in enumerate(rows_data, start=1):
        cells = table.add_row().cells

        no_surat = row.get("nomor_referensi", "")
        alamat = row.get("alamat", "")

        set_cell_text(
            cells[0],
            f"{idx}.",
            size=8,
            align=WD_ALIGN_PARAGRAPH.CENTER
        )

        set_cell_text(
            cells[1],
            no_surat,
            size=8,
            align=WD_ALIGN_PARAGRAPH.CENTER
        )

        set_cell_text(
            cells[2],
            alamat,
            size=7,
            align=WD_ALIGN_PARAGRAPH.LEFT
        )

        set_cell_text(
            cells[3],
            "",
            size=8,
            align=WD_ALIGN_PARAGRAPH.CENTER
        )

        set_cell_text(
            cells[4],
            "",
            size=8,
            align=WD_ALIGN_PARAGRAPH.CENTER
        )

        set_cell_text(
            cells[5],
            "",
            size=8,
            align=WD_ALIGN_PARAGRAPH.CENTER
        )

        set_cell_text(
            cells[6],
            "",
            size=8,
            align=WD_ALIGN_PARAGRAPH.CENTER
        )

    set_col_widths(table)

    doc.add_paragraph("")
    doc.add_paragraph("")

    # Tanda tangan
    ttd_table = doc.add_table(rows=1, cols=2)
    ttd_table.autofit = False

    left = ttd_table.cell(0, 0)
    right = ttd_table.cell(0, 1)

    left.width = Cm(12)
    right.width = Cm(12)

    left.text = ""
    right.text = ""

    # Kolom kiri: Pengirim
    p = left.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Pengirim,")
    set_font(r, size=9)

    for _ in range(3):
        left.add_paragraph("")

    p = left.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(data.get("petugas_nama", ""))
    set_font(r, size=9, bold=True)

    p = left.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(data.get("petugas_nip", ""))
    set_font(r, size=8)

    # Kolom kanan: Petugas Pos
    p = right.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"Jakarta, {data.get('tanggal_surat', '')}")
    set_font(r, size=9)

    p = right.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Petugas Pos")
    set_font(r, size=9)

    for _ in range(2):
        right.add_paragraph("")

    p = right.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(data.get("petugas_pos_nama", ""))
    set_font(r, size=9, bold=True)

    p = right.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(data.get("petugas_pos_nip", ""))
    set_font(r, size=8)

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)

    return bio
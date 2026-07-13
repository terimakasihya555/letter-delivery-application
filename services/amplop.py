from io import BytesIO

from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.section import WD_SECTION_START

from services.helpers import set_font


def set_paragraph_spacing(paragraph, before=0, after=0, line_spacing=1.0):
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)
    paragraph.paragraph_format.line_spacing = line_spacing


def add_blank_paragraph(doc, height_pt=12):
    p = doc.add_paragraph()
    set_paragraph_spacing(p, before=0, after=height_pt, line_spacing=1.0)
    r = p.add_run("")
    r.font.size = Pt(1)
    return p


def split_alamat(alamat):
    """
    Memecah alamat dari database menjadi beberapa baris.
    Baris pertama biasanya berisi:
    SDR. KETUA PENGADILAN ...
    """
    lines = []

    for line in str(alamat).splitlines():
        clean = line.strip()
        if clean:
            lines.append(clean)

    return lines


def add_envelope_page(doc, nomor_surat, alamat, is_first_page=False):
    """
    Membuat 1 halaman print amplop.

    Amplop sudah memiliki desain tercetak, sehingga file Word ini hanya
    menempatkan:
    1. Nomor surat pada area NOMOR
    2. Alamat tujuan pada area kanan amplop
    """
    if not is_first_page:
        doc.add_section(WD_SECTION_START.NEW_PAGE)

    section = doc.sections[-1]

    # Ukuran dibuat landscape mengikuti area amplop pada hasil scan.
    # Jika posisi print masih sedikit meleset, angka margin/spacing di bawah bisa disesuaikan.
    section.page_width = Cm(29.7)
    section.page_height = Cm(21.0)

    section.top_margin = Cm(0.7)
    section.bottom_margin = Cm(0.5)
    section.left_margin = Cm(1.0)
    section.right_margin = Cm(1.0)

    # =========================
    # POSISI NOMOR SURAT
    # =========================
    # Jarak atas untuk turun ke area tulisan "NOMOR"
    add_blank_paragraph(doc, height_pt=96)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_paragraph_spacing(p, before=0, after=0, line_spacing=1.0)

    # Indent ini menggeser nomor surat ke area titik-titik setelah "NOMOR :"
    p.paragraph_format.left_indent = Cm(3.05)

    r = p.add_run(nomor_surat)
    set_font(r, size=11, bold=False)

    # =========================
    # POSISI ALAMAT TUJUAN
    # =========================
    # Jarak dari nomor surat ke kotak alamat kanan.
    add_blank_paragraph(doc, height_pt=2)

    alamat_lines = split_alamat(alamat)

    if not alamat_lines:
        alamat_lines = [""]

    first_line = alamat_lines[0]
    next_lines = alamat_lines[1:]

    # Baris pertama:
    # "Kepada Yth." lebih kiri,
    # kemudian tab ke posisi "SDR...."
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_paragraph_spacing(p, before=0, after=0, line_spacing=1.0)

    # Posisi awal blok alamat kanan
    p.paragraph_format.left_indent = Cm(8.0)

    # Posisi tab untuk awal nama/alamat tujuan setelah "Kepada Yth."
    tab_position = Cm(3.1)
    p.paragraph_format.tab_stops.add_tab_stop(tab_position, WD_TAB_ALIGNMENT.LEFT)

    r = p.add_run("Kepada Yth.")
    set_font(r, size=11, bold=False)

    r = p.add_run("\t" + first_line)
    set_font(r, size=11, bold=False)

    # Baris berikutnya:
    # dibuat sejajar dengan awal "SDR...."
    for line in next_lines:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        set_paragraph_spacing(p, before=0, after=0, line_spacing=1.0)

        # 8.0 + 3.1 = sejajar dengan teks setelah tab pada baris pertama
        p.paragraph_format.left_indent = Cm(11.1)

        r = p.add_run(line)
        set_font(r, size=11, bold=False)


def build_amplop(data, kop_surat_path=None):
    """
    Membuat dokumen amplop berdasarkan surat yang dipilih dari database.

    Catatan:
    - kop_surat_path tidak dipakai lagi karena amplop fisik sudah memiliki kop/desain cetak.
    - setiap surat terpilih akan menjadi 1 halaman.
    """
    doc = Document()

    rows_data = data.get("selected_surat", [])

    for idx, row in enumerate(rows_data):
        nomor_surat = row.get("nomor_referensi", "")
        alamat = row.get("alamat", "")

        add_envelope_page(
            doc=doc,
            nomor_surat=nomor_surat,
            alamat=alamat,
            is_first_page=(idx == 0)
        )

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)

    return bio
from io import BytesIO
from itertools import count
from xml.sax.saxutils import escape

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.oxml import parse_xml
from docx.shared import Mm


# =========================================================
# KONFIGURASI UKURAN AMPLOP
# =========================================================

ENVELOPE_WIDTH_MM = 280
ENVELOPE_HEIGHT_MM = 110

# Posisi nomor surat
NOMOR_LEFT_MM = 70
NOMOR_BASELINE_FROM_BOTTOM_MM = 50

# Posisi alamat
ALAMAT_LEFT_MM = 100
ALAMAT_BASELINE_FROM_BOTTOM_MM = 40

# Font
FONT_NAME = "Times New Roman"
FONT_SIZE_PT = 11

# Jarak antarbaris alamat
LINE_SPACING_PT = 13

# Digunakan agar setiap textbox mempunyai ID berbeda
_TEXTBOX_COUNTER = count(1)


def mm_to_pt(mm_value):
    """
    Mengubah ukuran milimeter menjadi point.
    """
    return float(mm_value) * 72 / 25.4


def normalize_line(value):
    """
    Membersihkan spasi berlebih dalam satu baris.
    """
    if value is None:
        return ""

    return " ".join(str(value).split())


def split_alamat(alamat):
    """
    Membersihkan alamat dari database.

    Fungsi ini:
    - menyeragamkan jenis Enter;
    - menghapus baris kosong;
    - menghapus spasi berlebih;
    - mempertahankan urutan baris alamat.
    """
    if alamat is None:
        return []

    text = str(alamat)
    text = text.replace("\r\n", "\n")
    text = text.replace("\r", "\n")

    result = []

    for line in text.split("\n"):
        clean_line = normalize_line(line)

        if clean_line:
            result.append(clean_line)

    return result


def configure_envelope_section(section):
    """
    Mengatur halaman Word menjadi ukuran amplop 280 × 110 mm.
    Margin dibuat nol agar posisi dihitung dari sisi kertas.
    """
    section.page_width = Mm(ENVELOPE_WIDTH_MM)
    section.page_height = Mm(ENVELOPE_HEIGHT_MM)

    section.top_margin = Mm(0)
    section.bottom_margin = Mm(0)
    section.left_margin = Mm(0)
    section.right_margin = Mm(0)

    section.header_distance = Mm(0)
    section.footer_distance = Mm(0)


def add_absolute_textbox(
    paragraph,
    left_mm,
    baseline_from_bottom_mm,
    lines,
    width_mm,
    height_mm,
    font_size_pt=11,
    bold=True,
    line_spacing_pt=13,
):
    """
    Menambahkan textbox Word dengan posisi absolut.

    left_mm:
        Jarak awal tulisan dari sisi kiri kertas.

    baseline_from_bottom_mm:
        Jarak baseline baris pertama dari sisi bawah kertas.

    lines:
        Daftar baris yang akan dicetak.

    width_mm dan height_mm:
        Ukuran textbox.
    """
    if not lines:
        lines = [""]

    shape_id = next(_TEXTBOX_COUNTER)

    # Perkiraan jarak antara bagian atas kotak teks dan baseline huruf.
    # Nilai ini digunakan agar posisi baseline mendekati ukuran fisik.
    baseline_offset_mm = font_size_pt * 0.29

    top_mm = (
        ENVELOPE_HEIGHT_MM
        - baseline_from_bottom_mm
        - baseline_offset_mm
    )

    left_pt = mm_to_pt(left_mm)
    top_pt = mm_to_pt(top_mm)
    width_pt = mm_to_pt(width_mm)
    height_pt = mm_to_pt(height_mm)

    paragraph_xml = []

    for line in lines:
        safe_text = escape(normalize_line(line))
        bold_xml = "<w:b/>" if bold else ""

        paragraph_xml.append(
            f"""
            <w:p>
                <w:pPr>
                    <w:spacing
                        w:before="0"
                        w:after="0"
                        w:line="{int(line_spacing_pt * 20)}"
                        w:lineRule="exact"
                    />
                </w:pPr>

                <w:r>
                    <w:rPr>
                        <w:rFonts
                            w:ascii="{FONT_NAME}"
                            w:hAnsi="{FONT_NAME}"
                            w:eastAsia="{FONT_NAME}"
                        />
                        {bold_xml}
                        <w:sz w:val="{int(font_size_pt * 2)}"/>
                        <w:szCs w:val="{int(font_size_pt * 2)}"/>
                    </w:rPr>

                    <w:t xml:space="preserve">{safe_text}</w:t>
                </w:r>
            </w:p>
            """
        )

    textbox_xml = f"""
    <w:r
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:v="urn:schemas-microsoft-com:vml"
    >
        <w:pict>
            <v:rect
                id="EnvelopeTextBox{shape_id}"
                style="
                    position:absolute;
                    margin-left:{left_pt:.2f}pt;
                    margin-top:{top_pt:.2f}pt;
                    width:{width_pt:.2f}pt;
                    height:{height_pt:.2f}pt;
                    z-index:1;
                    mso-position-horizontal-relative:page;
                    mso-position-vertical-relative:page;
                    mso-wrap-style:none;
                "
                stroked="f"
                filled="f"
            >
                <v:textbox inset="0,0,0,0">
                    <w:txbxContent>
                        {''.join(paragraph_xml)}
                    </w:txbxContent>
                </v:textbox>
            </v:rect>
        </w:pict>
    </w:r>
    """

    paragraph._p.append(parse_xml(textbox_xml))


def add_envelope_page(
    doc,
    nomor_surat,
    alamat,
    is_first_page=False,
):
    """
    Membuat satu halaman amplop.

    Nomor surat:
    - mulai 70 mm dari kiri;
    - baseline 50 mm dari bawah;
    - bold.

    Alamat:
    - 'Kepada Yth.' mulai 100 mm dari kiri;
    - baseline 40 mm dari bawah;
    - alamat berikutnya berada di bawahnya;
    - bold.
    """
    if is_first_page:
        section = doc.sections[0]
    else:
        section = doc.add_section(WD_SECTION_START.NEW_PAGE)

    configure_envelope_section(section)

    canvas = doc.add_paragraph()
    canvas.paragraph_format.space_before = 0
    canvas.paragraph_format.space_after = 0
    canvas.paragraph_format.line_spacing = 1

    clean_nomor = normalize_line(nomor_surat)

    # =====================================================
    # NOMOR SURAT
    # =====================================================

    add_absolute_textbox(
        paragraph=canvas,
        left_mm=NOMOR_LEFT_MM,
        baseline_from_bottom_mm=NOMOR_BASELINE_FROM_BOTTOM_MM,
        lines=[clean_nomor],
        width_mm=105,
        height_mm=10,
        font_size_pt=FONT_SIZE_PT,
        bold=True,
        line_spacing_pt=LINE_SPACING_PT,
    )

    # =====================================================
    # ALAMAT PENERIMA
    # =====================================================

    alamat_lines = split_alamat(alamat)

    full_address_lines = ["Kepada Yth."] + alamat_lines

    add_absolute_textbox(
        paragraph=canvas,
        left_mm=ALAMAT_LEFT_MM,
        baseline_from_bottom_mm=ALAMAT_BASELINE_FROM_BOTTOM_MM,
        lines=full_address_lines,
        width_mm=170,
        height_mm=38,
        font_size_pt=FONT_SIZE_PT,
        bold=True,
        line_spacing_pt=LINE_SPACING_PT,
    )


def build_amplop(data, kop_surat_path=None):
    """
    Membuat dokumen amplop berdasarkan surat yang dipilih.

    Amplop fisik sudah memiliki kop dan desain tercetak.
    Dokumen Word hanya mencetak nomor surat dan alamat penerima.
    """
    doc = Document()

    # Mengatur section awal sebelum membuat halaman.
    configure_envelope_section(doc.sections[0])

    rows_data = data.get("selected_surat", [])

    for index, row in enumerate(rows_data):
        nomor_surat = row.get("nomor_referensi", "")
        alamat = row.get("alamat", "")

        add_envelope_page(
            doc=doc,
            nomor_surat=nomor_surat,
            alamat=alamat,
            is_first_page=(index == 0),
        )

    output = BytesIO()
    doc.save(output)
    output.seek(0)

    return output
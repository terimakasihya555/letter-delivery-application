import os
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, send_file, session

from services.db import (
    init_db,
    get_all_surat,
    get_surat_by_id,
    insert_surat,
    insert_many_surat,
    update_surat,
    delete_surat
)
from services.excel_importer import import_rekap_from_excel
from services.excel_exporter import build_rekap_excel
from services.surat_tanggung_jawab import build_surat_tanggung_jawab
from services.daftar_pengiriman import build_daftar_pengiriman
from services.amplop import build_amplop


app = Flask(__name__)
app.secret_key = "ganti_dengan_secret_key"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
KOP_SURAT_PATH = os.path.join(ASSETS_DIR, "kop_surat.docx")


def tanggal_hari_ini_indonesia():
    nama_bulan = {
        1: "Januari",
        2: "Februari",
        3: "Maret",
        4: "April",
        5: "Mei",
        6: "Juni",
        7: "Juli",
        8: "Agustus",
        9: "September",
        10: "Oktober",
        11: "November",
        12: "Desember",
    }

    hari_ini = datetime.now()
    return f"{hari_ini.day:02d} {nama_bulan[hari_ini.month]} {hari_ini.year}"


@app.before_request
def setup_database():
    init_db()


@app.route("/")
def home():
    return render_template("home.html")


# =========================
# HALAMAN REKAPITULASI
# =========================

@app.route("/rekap")
def rekap():
    q = request.args.get("q", "").strip()
    bulan = request.args.get("bulan", "").strip()
    tahun = request.args.get("tahun", "").strip()

    surat_list = get_all_surat(q=q, bulan=bulan, tahun=tahun)
    message = request.args.get("message", "")

    return render_template(
        "rekap.html",
        surat_list=surat_list,
        q=q,
        bulan=bulan,
        tahun=tahun,
        message=message
    )


@app.route("/rekap/export")
def export_rekap_excel():
    q = request.args.get("q", "").strip()
    bulan = request.args.get("bulan", "").strip()
    tahun = request.args.get("tahun", "").strip()

    surat_list = get_all_surat(q=q, bulan=bulan, tahun=tahun)

    fileobj = build_rekap_excel(
        surat_list=surat_list,
        q=q,
        bulan=bulan,
        tahun=tahun
    )

    if bulan and tahun:
        filename = f"rekapitulasi_surat_{bulan}_{tahun}.xlsx"
    elif tahun:
        filename = f"rekapitulasi_surat_{tahun}.xlsx"
    else:
        filename = "rekapitulasi_surat.xlsx"

    return send_file(
        fileobj,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/rekap/tambah", methods=["GET", "POST"])
def tambah_rekap():
    if request.method == "POST":
        data = request.form.to_dict()
        insert_surat(data)
        return redirect(url_for("rekap"))

    return render_template(
        "rekap_form.html",
        mode="tambah",
        surat=None
    )


@app.route("/rekap/edit/<int:surat_id>", methods=["GET", "POST"])
def edit_rekap(surat_id):
    surat = get_surat_by_id(surat_id)

    if not surat:
        return redirect(url_for("rekap"))

    if request.method == "POST":
        data = request.form.to_dict()
        update_surat(surat_id, data)
        return redirect(url_for("rekap"))

    return render_template(
        "rekap_form.html",
        mode="edit",
        surat=surat
    )


@app.route("/rekap/hapus/<int:surat_id>", methods=["POST"])
def hapus_rekap(surat_id):
    delete_surat(surat_id)
    return redirect(url_for("rekap"))


@app.route("/rekap/import", methods=["GET", "POST"])
def import_excel():
    if request.method == "POST":
        bulan = request.form.get("bulan", "").strip()
        tahun = request.form.get("tahun", "").strip()
        file = request.files.get("excel_file")

        if not bulan or not tahun:
            return render_template(
                "import_excel.html",
                error="Bulan dan tahun wajib diisi."
            )

        if not file or file.filename == "":
            return render_template(
                "import_excel.html",
                error="File Excel wajib diunggah."
            )

        if not file.filename.lower().endswith(".xlsx"):
            return render_template(
                "import_excel.html",
                error="File harus berformat .xlsx."
            )

        try:
            rows = import_rekap_from_excel(file, bulan, tahun)
            inserted_count = insert_many_surat(rows)

            return redirect(
                url_for(
                    "rekap",
                    bulan=bulan,
                    tahun=tahun,
                    message=f"Import berhasil. {inserted_count} data masuk ke database."
                )
            )

        except Exception as e:
            return render_template(
                "import_excel.html",
                error=f"Gagal import Excel: {e}"
            )

    return render_template("import_excel.html", error=None)


# =========================
# HALAMAN CETAK
# =========================

@app.route("/cetak", methods=["GET", "POST"])
def cetak():
    q = request.args.get("q", "").strip()
    bulan = request.args.get("bulan", "").strip()
    tahun = request.args.get("tahun", "").strip()

    surat_list = get_all_surat(q=q, bulan=bulan, tahun=tahun)

    if request.method == "POST":
        selected_ids = request.form.getlist("selected_surat")

        selected_rows = []

        for surat_id in selected_ids:
            surat = get_surat_by_id(int(surat_id))

            if surat:
                selected_rows.append(dict(surat))

        data = request.form.to_dict()

        data["tanggal_surat"] = tanggal_hari_ini_indonesia()
        data["kota"] = "Jakarta"

        data["selected_surat"] = selected_rows
        data["jumlah_dokumen"] = str(len(selected_rows))

        session["last_print_data"] = data

        return render_template(
            "hasil_cetak.html",
            jumlah=len(selected_rows)
        )

    return render_template(
        "cetak.html",
        surat_list=surat_list,
        q=q,
        bulan=bulan,
        tahun=tahun,
        kop_surat_terpasang=os.path.exists(KOP_SURAT_PATH)
    )


@app.route("/download/surat-tanggung-jawab", methods=["POST"])
def download_surat_tanggung_jawab():
    data = session.get("last_print_data", {})
    fileobj = build_surat_tanggung_jawab(data)

    return send_file(
        fileobj,
        as_attachment=True,
        download_name="surat_keterangan_tanggung_jawab.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


@app.route("/download/daftar-pengiriman", methods=["POST"])
def download_daftar_pengiriman():
    data = session.get("last_print_data", {})
    fileobj = build_daftar_pengiriman(data)

    return send_file(
        fileobj,
        as_attachment=True,
        download_name="daftar_pengiriman_pos_surat_dinas_x5.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


@app.route("/download/amplop", methods=["POST"])
def download_amplop():
    data = session.get("last_print_data", {})
    fileobj = build_amplop(data, kop_surat_path=KOP_SURAT_PATH)

    return send_file(
        fileobj,
        as_attachment=True,
        download_name="amplop_pengiriman.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


# =========================
# SETTING KOP SURAT AMPLOP
# =========================

@app.route("/settings/kop-surat", methods=["GET", "POST"])
def kop_surat_settings():
    message = None

    if request.method == "POST":
        file = request.files.get("kop_surat_file")

        if file and file.filename.lower().endswith(".docx"):
            os.makedirs(ASSETS_DIR, exist_ok=True)
            file.save(KOP_SURAT_PATH)
            message = "Kop surat berhasil disimpan dan akan otomatis dipakai untuk amplop."
        else:
            message = "File yang diunggah harus berformat .docx."

    return render_template(
        "kop_surat.html",
        message=message,
        is_set=os.path.exists(KOP_SURAT_PATH)
    )


@app.route("/settings/kop-surat/hapus", methods=["POST"])
def kop_surat_hapus():
    if os.path.exists(KOP_SURAT_PATH):
        os.remove(KOP_SURAT_PATH)

    return render_template(
        "kop_surat.html",
        message="Kop surat default telah dihapus.",
        is_set=False
    )


if __name__ == "__main__":
    app.run(debug=True)
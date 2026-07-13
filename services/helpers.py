from docx.shared import Pt


def set_font(run, size=12, bold=False, name="Times New Roman"):
    run.font.name = name
    run.font.size = Pt(size)
    run.bold = bold


def angka_ke_kata(n):
    satuan = [
        "",
        "satu",
        "dua",
        "tiga",
        "empat",
        "lima",
        "enam",
        "tujuh",
        "delapan",
        "sembilan"
    ]

    try:
        n = int(n)
    except (ValueError, TypeError):
        return str(n)

    if n == 0:
        return "nol"

    if n < 0:
        return "minus " + angka_ke_kata(abs(n))

    if n < 10:
        return satuan[n]

    if n < 20:
        if n == 10:
            return "sepuluh"
        if n == 11:
            return "sebelas"
        return f"{satuan[n - 10]} belas"

    if n < 100:
        sisa = n % 10
        return f"{satuan[n // 10]} puluh" + (f" {angka_ke_kata(sisa)}" if sisa else "")

    if n < 200:
        sisa = n - 100
        return "seratus" + (f" {angka_ke_kata(sisa)}" if sisa else "")

    if n < 1000:
        sisa = n % 100
        return f"{satuan[n // 100]} ratus" + (f" {angka_ke_kata(sisa)}" if sisa else "")

    if n < 2000:
        sisa = n - 1000
        return "seribu" + (f" {angka_ke_kata(sisa)}" if sisa else "")

    if n < 1000000:
        sisa = n % 1000
        return f"{angka_ke_kata(n // 1000)} ribu" + (f" {angka_ke_kata(sisa)}" if sisa else "")

    if n < 1000000000:
        sisa = n % 1000000
        return f"{angka_ke_kata(n // 1000000)} juta" + (f" {angka_ke_kata(sisa)}" if sisa else "")

    return str(n)
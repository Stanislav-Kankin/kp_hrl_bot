import os
from collections import OrderedDict
from docx.shared import Pt

import subprocess
import shutil
import tempfile


file_id_mapping = OrderedDict()


def clean_input(value):
    try:
        # Округляем до целого числа
        return int(round(float(value.replace(',', '.').strip())))
    except ValueError:
        raise ValueError(f"Некорректное значение: {value}")


def format_cost(value, with_ruble=False):
    text = f"{int(round(float(value))):,}".replace(',', ' ')
    return f"{text} ₽" if with_ruble else text


def format_count(value):
    return f"{int(value)}"


def cleanup_kp_files():
    current_dir = os.getcwd()
    for filename in os.listdir(current_dir):
        if filename.startswith(
            "КП_") and (
                filename.endswith(".docx") or filename.endswith(".pdf")):
            os.remove(os.path.join(current_dir, filename))


def set_montserrat_font(doc):
    styles = doc.styles
    for style in styles:
        if style.type == 1:
            font = style.font
            font.name = 'Montserrat'
            font.size = Pt(10)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    run = paragraph.runs[
                        0
                        ] if paragraph.runs else paragraph.add_run()
                    run.font.name = 'Montserrat'
                    run.font.size = Pt(10)


def convert_docx_to_pdf(docx_path: str) -> str | None:
    if not shutil.which("libreoffice") and not shutil.which("soffice"):
        raise RuntimeError("LibreOffice (libreoffice или soffice) не найдена в системе.")

    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"Файл не найден: {docx_path}")

    output_dir = tempfile.mkdtemp()

    try:
        subprocess.run(
            ["xvfb-run", "libreoffice", "--headless", "--convert-to", "pdf", docx_path, "--outdir", output_dir],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"Ошибка конвертации LibreOffice:\n{e.stderr.decode()}")

    base_name = os.path.splitext(os.path.basename(docx_path))[0]
    pdf_path = os.path.join(output_dir, f"{base_name}.pdf")

    # Убедимся, что PDF был создан
    if not os.path.exists(pdf_path):
        raise RuntimeError("Файл PDF не найден после конвертации.")

    return pdf_path

import os

from collections import OrderedDict

from docx.shared import Pt


file_id_mapping = OrderedDict()


def clean_input(value):
    try:
        return float(value.replace(',', '.').strip())
    except ValueError:
        raise ValueError(f"Некорректное значение: {value}")


def format_cost(value):
    return f"{value:,.2f}".replace(',', ' ').replace('.', ',')


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

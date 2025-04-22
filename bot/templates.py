from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from .utils import set_montserrat_font, format_cost, format_count


def load_template(template_name):
    return Document(f"templates/{template_name}")


def fill_standard_template(doc, data):
    table = doc.tables[0]
    set_montserrat_font(doc)

    # Заполняем количество сотрудников (жирным шрифтом)
    employee_count_cell = table.cell(1, 0)
    employee_count_cell.text = format_count(data["employee_license_count"])
    for paragraph in employee_count_cell.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in paragraph.runs:
            run.bold = True

    # Заполняем остальные ячейки с выравниванием по центру
    def fill_cell(row, col, text, bold=False):
        cell = table.cell(row, col)
        cell.text = text
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.bold = bold
                run.font.name = 'Montserrat'
                run.font.size = Pt(10)

    # Заполняем данные таблицы
    fill_cell(1, 2, format_cost(data["base_license_cost"]))
    fill_cell(1, 3, format_count(data["base_license_count"]))
    fill_cell(1, 5, format_cost(
        data["base_license_cost"] * data["base_license_count"]))

    fill_cell(2, 2, format_cost(data["hr_license_cost"]))
    fill_cell(2, 3, format_count(data["hr_license_count"]))
    fill_cell(2, 5, format_cost(
        data["hr_license_cost"] * data["hr_license_count"]))

    fill_cell(3, 2, format_cost(data["employee_license_cost"]))
    fill_cell(3, 3, format_count(data["employee_license_count"]))
    fill_cell(3, 5, format_cost(
        data["employee_license_cost"] * data["employee_license_count"]))

    if data["need_onprem"]:
        fill_cell(4, 2, format_cost(data["onprem_cost"]))
        fill_cell(4, 3, format_count(data["onprem_count"]))
        fill_cell(4, 4, "12")
        fill_cell(4, 5, format_cost(
            data["onprem_cost"] * data["onprem_count"]))
    else:
        fill_cell(4, 2, "-")
        fill_cell(4, 3, "-")
        fill_cell(4, 4, "-")
        fill_cell(4, 5, "-")

    total = (data["base_license_cost"] * data["base_license_count"] +
             data["hr_license_cost"] * data["hr_license_count"] +
             data["employee_license_cost"] * data["employee_license_count"])
    if data["need_onprem"]:
        total += data["onprem_cost"] * data["onprem_count"]

    fill_cell(5, 5, format_cost(total), bold=True)


def fill_complex_template(doc, data):
    set_montserrat_font(doc)
    company_name = data.get('company_name', '')

    for paragraph in doc.paragraphs:
        if "Коммерческое предложение HRlink для компании" in paragraph.text:
            paragraph.clear()
            run1 = paragraph.add_run(
                "Коммерческое предложение HRlink для компании "
                )
            run1.bold = True
            run1.font.size = Pt(18)
            run2 = paragraph.add_run(f'"{company_name}"')
            run2.bold = True
            run2.font.color.rgb = RGBColor(0x44, 0x9D, 0xE6)
            run2.font.size = Pt(18)
            break

    if len(doc.tables) > 0:
        table = doc.tables[0]

        # Аналогичное форматирование для комплексного шаблона
        def fill_cell(row, col, text, bold=False):
            cell = table.cell(row, col)
            cell.text = text
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.bold = bold
                    run.font.name = 'Montserrat'
                    run.font.size = Pt(10)

        fill_cell(1, 2, format_cost(data["base_license_cost"]))
        fill_cell(1, 3, format_count(data["base_license_count"]))
        fill_cell(1, 5, format_cost(
            data["base_license_cost"] * data["base_license_count"]))

        fill_cell(2, 2, format_cost(data["hr_license_cost"]))
        fill_cell(2, 3, format_count(data["hr_license_count"]))
        fill_cell(2, 5, format_cost(
            data["hr_license_cost"] * data["hr_license_count"]))

        fill_cell(3, 2, format_cost(data["employee_license_cost"]))
        fill_cell(3, 3, format_count(data["employee_license_count"]))
        fill_cell(3, 5, format_cost(
            data["employee_license_cost"] * data["employee_license_count"]))

        total = (data["base_license_cost"] * data["base_license_count"] +
                 data["hr_license_cost"] * data["hr_license_count"] +
                 data["employee_license_cost"] * data[
                     "employee_license_count"])
        fill_cell(4, 5, format_cost(total), bold=True)

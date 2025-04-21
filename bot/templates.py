from docx import Document
from docx.shared import Pt  # Добавляем импорт Pt
from .utils import set_montserrat_font, format_cost, format_count

def load_template(template_name):
    return Document(f"templates/{template_name}")

def fill_standard_template(doc, data):
    table = doc.tables[0]
    set_montserrat_font(doc)

    table.cell(1, 2).text = format_cost(data["base_license_cost"])
    table.cell(1, 3).text = format_count(data["base_license_count"])
    table.cell(1, 5).text = format_cost(data["base_license_cost"] * data["base_license_count"])

    table.cell(2, 2).text = format_cost(data["hr_license_cost"])
    table.cell(2, 3).text = format_count(data["hr_license_count"])
    table.cell(2, 5).text = format_cost(data["hr_license_cost"] * data["hr_license_count"])

    table.cell(3, 2).text = format_cost(data["employee_license_cost"])
    table.cell(3, 3).text = format_count(data["employee_license_count"])
    table.cell(3, 5).text = format_cost(data["employee_license_cost"] * data["employee_license_count"])

    if data["need_onprem"]:
        table.cell(4, 2).text = format_cost(data["onprem_cost"])
        table.cell(4, 3).text = format_count(data["onprem_count"])
        table.cell(4, 4).text = "12"
        table.cell(4, 5).text = format_cost(data["onprem_cost"] * data["onprem_count"])
    else:
        table.cell(4, 2).text = "-"
        table.cell(4, 3).text = "-"
        table.cell(4, 4).text = "-"
        table.cell(4, 5).text = "-"

    total = (data["base_license_cost"] * data["base_license_count"] +
             data["hr_license_cost"] * data["hr_license_count"] +
             data["employee_license_cost"] * data["employee_license_count"])
    if data["need_onprem"]:
        total += data["onprem_cost"] * data["onprem_count"]

    table.cell(5, 5).text = format_cost(total)

    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = 'Montserrat'
                    run.font.size = Pt(10)

def fill_complex_template(doc, data):
    set_montserrat_font(doc)

    # Заполняем название компании
    doc.paragraphs[0].text = f"Коммерческое предложение HRlink для компании “{data['company_name']}”."

    # Если в шаблоне есть таблица, заполняем её
    if len(doc.tables) > 0:
        table = doc.tables[0]

        table.cell(0, 2).text = format_cost(data["base_license_cost"])
        table.cell(0, 3).text = format_count(data["base_license_count"])
        table.cell(0, 5).text = format_cost(data["base_license_cost"] * data["base_license_count"])

        table.cell(1, 2).text = format_cost(data["hr_license_cost"])
        table.cell(1, 3).text = format_count(data["hr_license_count"])
        table.cell(1, 5).text = format_cost(data["hr_license_cost"] * data["hr_license_count"])

        table.cell(2, 2).text = format_cost(data["employee_license_cost"])
        table.cell(2, 3).text = format_count(data["employee_license_count"])
        table.cell(2, 5).text = format_cost(data["employee_license_cost"] * data["employee_license_count"])

        total = (data["base_license_cost"] * data["base_license_count"] +
                 data["hr_license_cost"] * data["hr_license_count"] +
                 data["employee_license_cost"] * data["employee_license_count"])

        table.cell(3, 5).text = format_cost(total)

        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = 'Montserrat'
                        run.font.size = Pt(10)

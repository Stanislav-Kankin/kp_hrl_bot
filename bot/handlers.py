from aiogram import Router, types, Bot
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.types import InlineKeyboardButton, InlineKeyboardMarkup
from bot.states import FormStandard, FormComplex, FormMarketing
from .utils import (
    clean_input, cleanup_kp_files, file_id_mapping,
    convert_to_pdf_libreoffice
    )
from .templates import (
    load_template, fill_standard_template, fill_complex_template,
    fill_marketing_template
    )
import uuid
from datetime import datetime
from io import BytesIO
from docx import Document
import re

router = Router()


@router.message(Command("start"))
async def start(message: types.Message):
    await message.answer(
        "Это бот для создания <b>КП</b>.\n"
        "Нажмите /kp для начала.\n"
        "Выбери актуальный шаблон КП.\n"
        "После его формирование можно будет выгрузить его в формате PDF.",
        )


@router.message(Command("kp"))
async def start_kp(message: types.Message, state: FSMContext):
    cleanup_kp_files()

    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(
            text="Стандартный шаблон HRL",
            callback_data="template_standard"
            )],
        [InlineKeyboardButton(
            text="HRL комплекс",
            callback_data="template_complex"
            )],
        [InlineKeyboardButton(
            text="Шаблон Маркетинг (общий)",
            callback_data="template_marketing"
            )]
    ])
    await message.answer("Какой шаблон КП интересует?", reply_markup=keyboard)


@router.callback_query(lambda c: c.data.startswith("template_"))
async def process_template_choice(callback: types.CallbackQuery,
                                  state: FSMContext):
    template_choice = callback.data.split("_")[1]
    await state.update_data(template_choice=template_choice)

    if template_choice == "complex":
        await state.set_state(FormComplex.company_name)
        await callback.message.answer("Введите <b>название компании</b>:")
    elif template_choice == "marketing":
        await state.set_state(FormMarketing.company_name)
        await callback.message.answer("Введите <b>название компании</b>:")
    else:
        await state.set_state(FormStandard.is_standard_pricing)
        keyboard = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(
                text="Да",
                callback_data="standard_pricing_yes")],
            [InlineKeyboardButton(
                text="Нет",
                callback_data="standard_pricing_no")]
        ])
        await callback.message.answer(
            "Стоимость Базовой лицензии и Лицензии "
            "Кадровика стандартная? (15 000,00 руб/год)",
            reply_markup=keyboard
        )
    await callback.answer()


@router.callback_query(lambda c: c.data.startswith("standard_pricing_"))
async def process_standard_pricing(callback: types.CallbackQuery,
                                   state: FSMContext):
    choice = callback.data.split("_")[2]
    await state.update_data(is_standard_pricing=(choice == "yes"))

    current_state = await state.get_state()
    if "FormStandard" in str(current_state):
        if choice == "yes":
            await state.update_data(
                base_license_cost=15000,
                hr_license_cost=15000
            )
            await state.set_state(FormStandard.base_license_count)
            await callback.message.answer(
                "Введите <b>количество Базовых лицензий</b>:"
                )
        else:
            await state.set_state(FormStandard.base_license_cost)
            await callback.message.answer(
                "Введите <b>стоимость Базовой лицензии</b> (руб/год):"
                )
    elif "FormComplex" in str(current_state):
        if choice == "yes":
            await state.update_data(
                base_license_cost=15000,
                hr_license_cost=15000
            )
            await state.set_state(FormComplex.base_license_count)
            await callback.message.answer(
                "Введите <b>количество Базовых лицензий</b>:"
                )
        else:
            await state.set_state(FormComplex.base_license_cost)
            await callback.message.answer(
                "Введите <b>стоимость Базовой лицензии</b> (руб/год):"
                )
    elif "FormMarketing" in str(current_state):
        if choice == "yes":
            await state.update_data(
                base_license_cost=15000,
                hr_license_cost=15000
            )
            await state.set_state(FormMarketing.base_license_count)
            await callback.message.answer(
                "Введите <b>количество Базовых лицензий</b>:"
                )
        else:
            await state.set_state(FormMarketing.base_license_cost)
            await callback.message.answer(
                "Введите <b>стоимость Базовой лицензии</b> (руб/год):"
                )

    await callback.answer()


@router.message(FormComplex.company_name)
async def process_company_name(message: types.Message, state: FSMContext):
    await state.update_data(company_name=message.text)
    await state.set_state(FormComplex.is_standard_pricing)
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(
            text="Да",
            callback_data="standard_pricing_yes")],
        [InlineKeyboardButton(
            text="Нет",
            callback_data="standard_pricing_no")]
    ])
    await message.answer(
        "Стоимость Базовой лицензии и Лицензии "
        "Кадровика стандартная? (15 000,00 руб/год)",
        reply_markup=keyboard
    )


@router.message(FormMarketing.company_name)
async def process_marketing_company_name(message: types.Message, state: FSMContext):
    await state.update_data(company_name=message.text)
    await state.set_state(FormMarketing.is_standard_pricing)
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(
            text="Да",
            callback_data="standard_pricing_yes")],
        [InlineKeyboardButton(
            text="Нет",
            callback_data="standard_pricing_no")]
    ])
    await message.answer(
        "Стоимость Базовой лицензии и Лицензии "
        "Кадровика стандартная? (15 000,00 руб/год)",
        reply_markup=keyboard
    )


@router.message(FormStandard.base_license_cost)
async def process_base_license_cost_standard(message: types.Message,
                                             state: FSMContext):
    try:
        value = clean_input(message.text)
        await state.update_data(base_license_cost=value)
        await state.set_state(FormStandard.base_license_count)
        await message.answer("Введите <b>количество Базовых лицензий</b>:")
    except ValueError as e:
        await message.answer(
            f"Ошибка: {str(e)}. Пожалуйста, введите корректное значение."
            )


@router.message(FormComplex.base_license_cost)
async def process_base_license_cost_complex(message: types.Message,
                                            state: FSMContext):
    try:
        value = clean_input(message.text)
        await state.update_data(base_license_cost=value)
        await state.set_state(FormComplex.base_license_count)
        await message.answer(
            "Введите <b>количество Базовых лицензий</b>:"
            )
    except ValueError as e:
        await message.answer(
            f"Ошибка: {str(e)}. Пожалуйста, введите корректное значение."
            )


@router.message(FormMarketing.base_license_cost)
async def process_base_license_cost_marketing(message: types.Message,
                                             state: FSMContext):
    try:
        value = clean_input(message.text)
        await state.update_data(base_license_cost=value)
        await state.set_state(FormMarketing.base_license_count)
        await message.answer("Введите <b>количество Базовых лицензий</b>:")
    except ValueError as e:
        await message.answer(
            f"Ошибка: {str(e)}. Пожалуйста, введите корректное значение."
            )


@router.message(FormStandard.base_license_count)
async def process_base_license_count_standard(message: types.Message,
                                              state: FSMContext):
    try:
        value = clean_input(message.text)
        await state.update_data(base_license_count=value)

        data = await state.get_data()
        if data.get("is_standard_pricing", False):
            await state.set_state(FormStandard.hr_license_count)
            await message.answer(
                "Введите <b>количество лицензий кадровиков</b>:"
                )
        else:
            await state.set_state(FormStandard.hr_license_cost)
            await message.answer(
                "Введите <b>стоимость лицензий кадровиков</b> (руб/год):"
                )

    except ValueError as e:
        await message.answer(
            f"Ошибка: {str(e)}. Пожалуйста, введите корректное значение."
            )


@router.message(FormComplex.base_license_count)
async def process_base_license_count_complex(message: types.Message,
                                             state: FSMContext):
    try:
        value = clean_input(message.text)
        await state.update_data(base_license_count=value)

        data = await state.get_data()
        if data.get("is_standard_pricing", False):
            await state.set_state(FormComplex.hr_license_count)
            await message.answer(
                "Введите <b>количество лицензий кадровиков</b>:"
                )
        else:
            await state.set_state(FormComplex.hr_license_cost)
            await message.answer(
                "Введите <b>стоимость лицензий кадровиков</b> (руб/год):"
                )

    except ValueError as e:
        await message.answer(
            f"Ошибка: {str(e)}. Пожалуйста, введите корректное значение."
            )


@router.message(FormMarketing.base_license_count)
async def process_base_license_count_marketing(message: types.Message,
                                              state: FSMContext):
    try:
        value = clean_input(message.text)
        await state.update_data(base_license_count=value)

        data = await state.get_data()
        if data.get("is_standard_pricing", False):
            await state.set_state(FormMarketing.hr_license_count)
            await message.answer(
                "Введите <b>количество лицензий кадровиков</b>:"
                )
        else:
            await state.set_state(FormMarketing.hr_license_cost)
            await message.answer(
                "Введите <b>стоимость лицензий кадровиков</b> (руб/год):"
                )

    except ValueError as e:
        await message.answer(
            f"Ошибка: {str(e)}. Пожалуйста, введите корректное значение."
            )


@router.message(FormStandard.hr_license_cost)
async def process_hr_license_cost_standard(message: types.Message,
                                           state: FSMContext):
    try:
        value = clean_input(message.text)
        await state.update_data(hr_license_cost=value)
        await state.set_state(FormStandard.hr_license_count)
        await message.answer(
            "Введите <b>количество лицензий кадровиков</b>:"
            )
    except ValueError as e:
        await message.answer(
            f"Ошибка: {str(e)}. Пожалуйста, введите корректное значение."
            )


@router.message(FormComplex.hr_license_cost)
async def process_hr_license_cost_complex(message: types.Message,
                                          state: FSMContext):
    try:
        value = clean_input(message.text)
        await state.update_data(hr_license_cost=value)
        await state.set_state(FormComplex.hr_license_count)
        await message.answer(
            "Введите <b>количество лицензий кадровиков</b>:")
    except ValueError as e:
        await message.answer(
            f"Ошибка: {str(e)}. Пожалуйста, введите корректное значение."
            )


@router.message(FormMarketing.hr_license_cost)
async def process_hr_license_cost_marketing(message: types.Message,
                                           state: FSMContext):
    try:
        value = clean_input(message.text)
        await state.update_data(hr_license_cost=value)
        await state.set_state(FormMarketing.hr_license_count)
        await message.answer(
            "Введите <b>количество лицензий кадровиков</b>:"
            )
    except ValueError as e:
        await message.answer(
            f"Ошибка: {str(e)}. Пожалуйста, введите корректное значение."
            )


@router.message(FormStandard.hr_license_count)
async def process_hr_license_count_standard(message: types.Message,
                                            state: FSMContext):
    try:
        value = clean_input(message.text)
        await state.update_data(hr_license_count=value)
        await state.set_state(FormStandard.employee_license_cost)
        await message.answer(
            "Введите <b>стоимость лицензии сотрудника</b> (руб/год):"
            )
    except ValueError as e:
        await message.answer(
            f"Ошибка: {str(e)}. Пожалуйста, введите корректное значение."
            )


@router.message(FormComplex.hr_license_count)
async def process_hr_license_count_complex(message: types.Message,
                                           state: FSMContext):
    try:
        value = clean_input(message.text)
        await state.update_data(hr_license_count=value)
        await state.set_state(FormComplex.employee_license_cost)
        await message.answer(
            "Введите <b>стоимость лицензии сотрудника</b> (руб/год):")

    except ValueError as e:
        await message.answer(
            f"Ошибка: {str(e)}. Пожалуйста, введите корректное значение."
            )


@router.message(FormMarketing.hr_license_count)
async def process_hr_license_count_marketing(message: types.Message,
                                            state: FSMContext):
    try:
        value = clean_input(message.text)
        await state.update_data(hr_license_count=value)
        await state.set_state(FormMarketing.employee_license_cost)
        await message.answer(
            "Введите <b>стоимость лицензии сотрудника</b> (руб/год):"
            )
    except ValueError as e:
        await message.answer(
            f"Ошибка: {str(e)}. Пожалуйста, введите корректное значение."
            )


@router.message(FormStandard.employee_license_cost)
async def process_employee_license_cost_standard(message: types.Message,
                                                 state: FSMContext):
    try:
        value = clean_input(message.text)
        await state.update_data(employee_license_cost=value)
        await state.set_state(FormStandard.employee_license_count)
        await message.answer(
            "Введите <b>количество лицензий сотрудника</b>:"
            )
    except ValueError as e:
        await message.answer(
            f"Ошибка: {str(e)}. Пожалуйста, введите корректное значение."
            )


@router.message(FormComplex.employee_license_cost)
async def process_employee_license_cost_complex(message: types.Message,
                                                state: FSMContext):
    try:
        value = clean_input(message.text)
        await state.update_data(employee_license_cost=value)
        await state.set_state(FormComplex.employee_license_count)
        await message.answer(
            "Введите <b>количество лицензий сотрудника</b>:"
            )
    except ValueError as e:
        await message.answer(
            f"Ошибка: {str(e)}. Пожалуйста, введите корректное значение."
            )


@router.message(FormMarketing.employee_license_cost)
async def process_employee_license_cost_marketing(message: types.Message,
                                                 state: FSMContext):
    try:
        value = clean_input(message.text)
        await state.update_data(employee_license_cost=value)
        await state.set_state(FormMarketing.employee_license_count)
        await message.answer(
            "Введите <b>количество лицензий сотрудника</b>:"
            )
    except ValueError as e:
        await message.answer(
            f"Ошибка: {str(e)}. Пожалуйста, введите корректное значение."
            )


@router.message(FormStandard.employee_license_count)
async def process_employee_license_count_standard(message: types.Message,
                                                  state: FSMContext):
    try:
        value = clean_input(message.text)
        await state.update_data(employee_license_count=value)

        keyboard = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="Да", callback_data="onprem_yes")],
            [InlineKeyboardButton(text="Нет", callback_data="onprem_no")]
        ])
        await message.answer(
            "Нужен ли on-prem?", reply_markup=keyboard)
        await state.set_state(FormStandard.need_onprem)
    except ValueError as e:
        await message.answer(
            f"Ошибка: {str(e)}. Пожалуйста, введите корректное значение."
            )


@router.message(FormMarketing.employee_license_count)
async def process_employee_license_count_marketing(message: types.Message,
                                                  state: FSMContext):
    try:
        value = clean_input(message.text)
        await state.update_data(employee_license_count=value)

        keyboard = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="Да", callback_data="onprem_yes")],
            [InlineKeyboardButton(text="Нет", callback_data="onprem_no")]
        ])
        await message.answer(
            "Нужен ли on-prem?", reply_markup=keyboard)
        await state.set_state(FormMarketing.need_onprem)
    except ValueError as e:
        await message.answer(
            f"Ошибка: {str(e)}. Пожалуйста, введите корректное значение."
            )


@router.message(FormComplex.employee_license_count)
async def process_employee_license_count_complex(message: types.Message,
                                                 state: FSMContext):
    try:
        value = clean_input(message.text)
        await state.update_data(employee_license_count=value)
        await state.set_state(FormComplex.kp_expiration)
        await message.answer("Введите <b>срок действия КП</b> в формате дд.мм.гггг:")
    except ValueError as e:
        await message.answer(
            f"Ошибка: {str(e)}. Пожалуйста, введите корректное значение.")


@router.callback_query(lambda c: c.data.startswith("onprem_"))
async def process_onprem_choice(callback: types.CallbackQuery,
                                state: FSMContext):
    choice = callback.data.split("_")[1]
    current_state = await state.get_state()

    if choice == "yes":
        await state.update_data(need_onprem=True)
        await state.set_state(FormStandard.onprem_cost)
        await callback.message.answer("Введите <b>сумму on-prem</b> (руб/год):")
    else:
        await state.update_data(
            need_onprem=False,
            onprem_cost=0,
            onprem_count=0
        )

        # ⬇️ Вместо генерации — спросим дату
        if "FormStandard" in str(current_state):
            await state.set_state(FormStandard.kp_expiration)
        elif "FormMarketing" in str(current_state):
            await state.set_state(FormMarketing.kp_expiration)
        elif "FormComplex" in str(current_state):
            await state.set_state(FormComplex.kp_expiration)

        await callback.message.answer("Введите <b>срок действия КП</b> в формате дд.мм.гггг:")

    await callback.answer()


@router.message(FormStandard.onprem_cost)
async def process_onprem_cost_standard(message: types.Message, state: FSMContext):
    try:
        value = clean_input(message.text)
        await state.update_data(onprem_cost=value)
        await state.set_state(FormStandard.onprem_count)
        await message.answer("Введите <b>количество лицензий on-prem</b>:")
    except ValueError as e:
        await message.answer(
            f"Ошибка: {str(e)}. Пожалуйста, введите корректное значение.")


@router.message(FormMarketing.onprem_cost)
async def process_onprem_cost_marketing(message: types.Message, state: FSMContext):
    try:
        value = clean_input(message.text)
        await state.update_data(onprem_cost=value)
        await state.set_state(FormMarketing.onprem_count)
        await message.answer("Введите <b>количество лицензий on-prem</b>:")
    except ValueError as e:
        await message.answer(
            f"Ошибка: {str(e)}. Пожалуйста, введите корректное значение.")


@router.message(FormStandard.onprem_count)
async def process_onprem_count_standard(message: types.Message, state: FSMContext):
    try:
        value = clean_input(message.text)
        await state.update_data(onprem_count=value)
        await state.set_state(FormStandard.kp_expiration)
        await message.answer("Введите <b>срок действия КП</b> в формате дд.мм.гггг:")
    except ValueError as e:
        await message.answer(
            f"Ошибка: {str(e)}. Пожалуйста, введите корректное значение.")


@router.message(FormMarketing.onprem_count)
async def process_onprem_count_marketing(message: types.Message, state: FSMContext):
    try:
        value = clean_input(message.text)
        await state.update_data(onprem_count=value)
        await state.set_state(FormMarketing.kp_expiration)
        await message.answer("Введите <b>срок действия КП</b> в формате дд.мм.гггг:")
    except ValueError as e:
        await message.answer(
            f"Ошибка: {str(e)}. Пожалуйста, введите корректное значение.")


async def generate_kp(bot: Bot, message: types.Message, state: FSMContext):
    data = await state.get_data()
    template_choice = data.get("template_choice", "standard")

    if data.get("is_standard_pricing", False):
        if "base_license_cost" not in data:
            data["base_license_cost"] = 15000
        if "hr_license_cost" not in data:
            data["hr_license_cost"] = 15000

    if template_choice == "standard":
        doc = load_template("template.docx", need_onprem=data.get("need_onprem", True))
        fill_standard_template(doc, data)
    elif template_choice == "complex":
        doc = load_template("template_complex.docx")
        fill_complex_template(doc, data)
    elif template_choice == "marketing":
        doc = Document(f"templates/{'template_.docx' if data.get('need_onprem', False) else 'template_m_no.docx'}")

        fill_marketing_template(doc, data)

    kp_filename = f"КП_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    doc.save(kp_filename)

    with open(kp_filename, 'rb') as file:
        doc_message = await message.answer_document(
            types.BufferedInputFile(file.read(), filename=kp_filename),
            caption="Ваше КП готово!\n"
            "Вы можете скачать его или конвертировать в PDF."
        )

    unique_id = str(uuid.uuid4())

    if len(file_id_mapping) >= 5:
        oldest_file_id = next(iter(file_id_mapping))
        del file_id_mapping[oldest_file_id]

    file_id_mapping[unique_id] = doc_message.document.file_id

    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(
            text="Сделать PDF",
            callback_data=f"convert_to_pdf_{unique_id}")]
    ])
    await message.answer(
        "Нажми ниже для формирования PDF",
        reply_markup=keyboard)

    await state.clear()



@router.callback_query(lambda c: c.data.startswith("convert_to_pdf_"))
async def convert_to_pdf(callback: types.CallbackQuery, bot: Bot):
    unique_id = callback.data.split("_")[3]
    file_id = file_id_mapping.get(unique_id)

    if not file_id:
        await callback.message.answer("Файл не найден.")
        await callback.answer()
        return

    file_info = await bot.get_file(file_id)
    file_path = file_info.file_path

    file = await bot.download_file(file_path)
    docx_path = f"temp_{unique_id}.docx"
    with open(docx_path, "wb") as f:
        f.write(file.read())

    try:
        pdf_path = convert_to_pdf_libreoffice(docx_path)

        with open(pdf_path, "rb") as f:
            await callback.message.answer_document(
                types.BufferedInputFile(f.read(), filename=os.path.basename(pdf_path)),
                caption="Ваше КП в формате PDF."
            )
    except Exception as e:
        await callback.message.answer(f"Ошибка при конвертации: {e}")
    finally:
        # Удалим временные файлы
        for path in (docx_path, pdf_path):
            if os.path.exists(path):
                os.remove(path)

    await callback.answer()


@router.message(FormStandard.kp_expiration)
@router.message(FormMarketing.kp_expiration)
@router.message(FormComplex.kp_expiration)
async def process_kp_expiration(message: types.Message, state: FSMContext):
    date_text = message.text.strip()

    if not re.match(r"^\d{2}\.\d{2}\.\d{4}$", date_text):
        await message.answer("⛔️ Пожалуйста, введите дату в формате <b>дд.мм.гггг</b>, например: 30.06.2025")
        return

    await state.update_data(kp_expiration=date_text)
    await generate_kp(message.bot, message, state)
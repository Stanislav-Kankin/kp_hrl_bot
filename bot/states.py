from aiogram.fsm.state import State, StatesGroup


class FormStandard(StatesGroup):
    is_standard_pricing = State()
    base_license_cost = State()
    base_license_count = State()
    hr_license_cost = State()
    hr_license_count = State()
    employee_license_cost = State()
    employee_license_count = State()
    need_onprem = State()
    onprem_cost = State()
    onprem_count = State()
    kp_expiration = State()


class FormComplex(StatesGroup):
    company_name = State()
    is_standard_pricing = State()
    base_license_cost = State()
    base_license_count = State()
    hr_license_cost = State()
    hr_license_count = State()
    employee_license_cost = State()
    employee_license_count = State()
    kp_expiration = State()


class FormMarketing(StatesGroup):
    company_name = State()
    is_standard_pricing = State()
    base_license_cost = State()
    base_license_count = State()
    hr_license_cost = State()
    hr_license_count = State()
    employee_license_cost = State()
    employee_license_count = State()
    need_onprem = State()
    onprem_cost = State()
    onprem_count = State()
    kp_expiration = State()
    unep_count = State()
    sms_count = State()
    custom_conditions = State()
    current_condition = State()
    kp_expiration = State()

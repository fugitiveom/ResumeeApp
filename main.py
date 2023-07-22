''' it's a start file'''
import os
from datetime import date
from modules.usecases import WindowsWordCase, UseCaseDataDTO
from config import COUNTRY, JOB_TYPE, ROOT_DIR, JOB_PORTAL, PH_POSITION_TITLE, \
      PH_COMPANY_NAME, PH_DATE, PH_PLATFORM_SOURCE

company = input('Company: ')
position = input('Position: ')

WORKDIR = ROOT_DIR + COUNTRY

replacements = {
    PH_POSITION_TITLE: position,
    PH_COMPANY_NAME: company,
    PH_PLATFORM_SOURCE: JOB_PORTAL,
    PH_DATE: str(date.today())
}

use_case_data_dto = UseCaseDataDTO(company, JOB_TYPE, position, JOB_PORTAL)

if os.name == 'nt':
    winword_case = WindowsWordCase(WORKDIR, use_case_data_dto)
    winword_case.make_documents()

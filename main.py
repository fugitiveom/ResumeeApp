''' it's a start file'''
import os
from modules.usecases import WindowsWordCase, UseCaseDataDTO
from config import COUNTRY, JOB_TYPE, ROOT_DIR, JOB_PORTAL

company = input('Company: ')
position = input('Position: ')

WORKDIR = ROOT_DIR + COUNTRY

use_case_data_dto = UseCaseDataDTO(company, JOB_TYPE, position, JOB_PORTAL)

if os.name == 'nt':
    winword_case = WindowsWordCase(WORKDIR, use_case_data_dto)
    winword_case.make_documents()

''' it's a start file'''
import os
import sys
from datetime import date
import psutil
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

use_case_data_dto = UseCaseDataDTO(company, JOB_TYPE, position, JOB_PORTAL, replacements)

def check_preconditions() -> dict:
    ''' check all preconditions here '''
    if_already_sent = os.path.isdir(os.path.join(WORKDIR, '_Sent', company))
    if_path_exists = os.path.isdir(os.path.join(WORKDIR, company))
    word_processes = [proc.name() for proc in psutil.process_iter() \
                        if proc.name() == 'WINWORD.EXE']
    return {'if_already_sent': if_already_sent, 'word_processes': word_processes, \
                'if_path_exists': if_path_exists}

preconditions = check_preconditions()
if preconditions['if_path_exists']:
    ifcontinue = input('Каталог клиента уже существует, продолжить работу? (y/n): ')
    if ifcontinue != 'y':
        sys.exit()
if preconditions['if_already_sent']:
    ifcontinue = input('Вы уже отправляли резюме этой компании, продолжить? (y/n): ')
    if ifcontinue != 'y':
        sys.exit()
if preconditions['word_processes'] == ['WINWORD.EXE']:
    ifcontinue = input('Microsoft Word запущен, в случае продолжения он будет закрыт. '
                        'Продолжить? (y/n): ')
    if ifcontinue != 'y':
        sys.exit()

if os.name == 'nt':
    winword_case = WindowsWordCase(WORKDIR, use_case_data_dto)
    winword_case.make_documents()

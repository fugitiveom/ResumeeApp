''' it's a start file'''
import os
from modules.usecases import WindowsWordCase
from config import COUNTRY, JOB_TYPE, ROOT_DIR, JOB_PORTAL

company = input('Company: ')
position = input('Position: ')

WORKDIR = ROOT_DIR + COUNTRY

if os.name == 'nt':
    winword_case = WindowsWordCase(WORKDIR, company, JOB_TYPE, position, JOB_PORTAL)
    winword_case.make_documents()

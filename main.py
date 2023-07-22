''' it's a start file'''
import os
from modules.usecases import WindowsWordCase
from config import COUNTRY, JOB_TYPE, ROOT_DIR, JOB_PORTAL

company = input('Company: ')
position = input('Position: ')

WORKDIR = ROOT_DIR + COUNTRY

class UseCaseDataDTO:
    def __init__(self, company, job_type, position, job_portal):
        self.company = company
        self.job_type = job_type
        self.position = position
        self.job_portal = job_portal

if os.name == 'nt':
    winword_case = WindowsWordCase(WORKDIR, company, JOB_TYPE, position, JOB_PORTAL)
    winword_case.make_documents()

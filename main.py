from modules.usecases import *
import os
from config import country, job_type, root_dir, job_portal

company = input('Company: ')
position = input('Position: ')

workdir = root_dir + country

if os.name == 'nt':
    winword_case = WindowsWord_case(workdir, company, job_type, position, job_portal)
    winword_case.make_documents()

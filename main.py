import classes
import os
from config import country, job_type, root_dir, job_portal

company = input('Company: ')
position = input('Position: ')

workdir = root_dir + country

if os.name == 'nt':
    wingenerator = classes.GenerateDocsWin()
    wingenerator.generate(workdir, company, job_type, position, job_portal)

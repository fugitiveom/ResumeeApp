import modules
from config import country, job_type, root_dir, job_portal

company = input('Company: ')
position = input('Position: ')

workdir = root_dir + country

modules.makedir(workdir, company)
modules.copy_templates(workdir, company, job_type)
modules.generate_email(workdir, company, position,job_portal)
modules.generate_resume(workdir, company)
modules.generate_cover_letter(workdir, company, position, job_portal)
modules.remove_docx(workdir, company)
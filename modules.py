import os
import shutil
import glob
from config import email_pref


def makedir(workdir, company):
    if not os.path.isdir(workdir + company):
        os.mkdir(workdir + company)


def copy_templates(workdir, company, job_type):  # TODO убрать конкатенацию
    templates_path = workdir + '/_templates'
    templates = glob.glob(templates_path + '/*.docx')
    templates += glob.glob(templates_path + '/*.txt')
    for template in templates:
        if job_type == 't' and template.find('tech') != -1:
            shutil.copy(template, workdir + '/' + company)
        elif job_type == 'm' and template.find('manager') != -1:
            shutil.copy(template, workdir + '/' + company)


def generate_email(workdir, company, position, job_portal):  # TODO Переписать на построчный вариант
    source = glob.glob(workdir + '/' + company + '/' + email_pref)
    with open(source[0], 'r+') as f:
        data = f.read()
        data = data.replace('[position name]', position)
        data = data.replace('[Company Name]', company)
        data = data.replace('[Job Source]', job_portal)
        f.seek(0)
        f.write(data)


def generate_resume(workdir, company):
    source = glob.glob(workdir + '/' + company + '/' + email_pref)
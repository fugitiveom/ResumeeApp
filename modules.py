import os
import shutil
import glob
import sys

from config import email_regexp, resume_regexp, cover_letter_regexp
import win32com.client as win32
from datetime import date


def makedir(workdir, company):
    company_dir = os.path.join(workdir, company)
    company_sent_dir = os.path.join(workdir, '_Sent', company)
    if os.path.isdir(company_sent_dir):
        ifcontinue = input('Вы уже отправляли резюме этой компании, продолжить? (y/n): ')
        if ifcontinue != 'y':
            sys.exit()
    if not os.path.isdir(company_dir):
        os.makedirs(company_dir)


def copy_templates(workdir, company, job_type):  # TODO убрать конкатенацию
    templates_path = os.path.join(workdir + '/_templates')
    templates = glob.glob(templates_path + '/*.docx')
    templates += glob.glob(templates_path + '/*.txt')
    for template in templates:
        if job_type == 't' and template.find('tech') != -1:
            shutil.copy(template, workdir + '/' + company)
        elif job_type == 'm' and template.find('manager') != -1:
            shutil.copy(template, workdir + '/' + company)


def generate_email(workdir, company, position, job_portal):  # TODO Переписать на построчный вариант
    source = glob.glob(workdir + '/' + company + '/' + email_regexp)
    source[0] = source[0].replace('/', '\\')
    with open(source[0], 'r+') as f:
        data = f.read()
        data = data.replace('[position name]', position)
        data = data.replace('[Company Name]', company)
        data = data.replace('[Job Source]', job_portal)
        f.seek(0)
        f.write(data)


def generate_resume(workdir, company):
    source = glob.glob(workdir + '/' + company + '/' + resume_regexp)
    source[0] = source[0].replace('/', '\\')

    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False

    doc = word.Documents.Open(source[0])
    doc.SaveAs(source[0].replace('docx', 'pdf').replace('tech', company), 17)
    doc.Close()
    word.Quit()


def generate_cover_letter(workdir, company, position, job_portal):
    source = glob.glob(workdir + '/' + company + '/' + cover_letter_regexp)
    source[0] = source[0].replace('/', '\\')
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False

    replacements = {
        '[Position Title]' : position,
        '[Company Name]' : company,
        '[Platform/Source]' : job_portal,
        '[Date]' : str(date.today())
    }

    doc = word.Documents.Open(source[0])

    for find_text, replace_with in replacements.items():
        for paragraph in doc.Paragraphs:
            if find_text in paragraph.Range.Text:
                paragraph.Range.HighlightColorIndex = 0
                paragraph.Range.Text = paragraph.Range.Text.replace(find_text, replace_with)

    doc.SaveAs(source[0].replace('docx', 'pdf').replace('tech', company), 17)
    doc.Close()
    word.Quit()


def remove_docx(workdir, company):
    docxs = glob.glob(workdir + '/' + company + '/' + '*.docx')
    for docx in docxs:
        os.remove(docx)


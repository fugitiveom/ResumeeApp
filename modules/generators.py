import os
import shutil
import glob
import sys

import psutil

from config import email_regexp, resume_regexp, cover_letter_regexp
import win32com.client as win32
from datetime import date

class Preparator:
    def __init__(self, workdir, company, job_type):
        self.workdir = workdir
        self.company = company
        self.job_type = job_type

    def prepare_dir(self):
        self.makedir()
        self.copy_templates()

    def makedir(self, workdir, company):
        company_dir = os.path.join(workdir, company)
        company_sent_dir = os.path.join(workdir, '_Sent', company)
        if os.path.isdir(company_sent_dir):
            ifcontinue = input('Вы уже отправляли резюме этой компании, продолжить? (y/n): ')
            if ifcontinue != 'y':
                sys.exit()
        if not os.path.isdir(company_dir):
            os.makedirs(company_dir)

    def copy_templates(self, workdir, company, job_type):  # TODO убрать конкатенацию
        templates_path = os.path.join(workdir + '/_templates')
        templates = glob.glob(templates_path + '/*.docx')
        templates += glob.glob(templates_path + '/*.txt')
        for template in templates:
            if job_type == 't' and template.find('tech') != -1:
                shutil.copy(template, workdir + '/' + company)
            elif job_type == 'm' and template.find('manager') != -1:
                shutil.copy(template, workdir + '/' + company)

class WinWord:
    def __init__(self):
        self.word = None

    def check_word(self):
        for proc in psutil.process_iter():
            if proc.name() == 'WINWORD.EXE':
                term_agree = input('Microsoft Word запущен. Закрыть? (y/n): ')
                if term_agree == 'y':
                    proc.terminate()
    
    def open_word(self):
        self.check_word()
        self.word = win32.gencache.EnsureDispatch('Word.Application')
        self.word.Visible = False


    def close_word(self):
        self.word.Quit()

class Clear:
    def remove_docx(self, workdir, company):
        docxs = glob.glob(workdir + '/' + company + '/' + '*.docx')
        for docx in docxs:
            os.remove(docx)


class WinDocsGenerator(Preparator, Clear, WinWord):
    def __init__(self):
        super().__init__()
        self.prepare = Preparator()
        self.clear = Clear()
        self.winword = WinWord()

    def generate(self, workdir, company, job_type, position, job_portal):
        self.prepare.makedir(workdir, company)
        self.prepare.copy_templates(workdir, company, job_type)

        self.winword.open_word()
        
        self.generate_email(workdir, company, position, job_portal)
        self.generate_resume(workdir, company)
        self.generate_cover_letter(workdir, company, position, job_portal)

        self.winword.close_word()

        self.clear.remove_docx(workdir, company)
        

    def generate_email(self, workdir, company, position, job_portal):  # TODO Переписать на построчный вариант
        source = glob.glob(workdir + '/' + company + '/' + email_regexp)
        source[0] = source[0].replace('/', '\\')
        with open(source[0], 'r+') as f:
            data = f.read()
            data = data.replace('[position name]', position)
            data = data.replace('[Company Name]', company)
            data = data.replace('[Job Source]', job_portal)
            f.seek(0)
            f.write(data)
            f.truncate()

    def generate_resume(self, workdir, company):
        source = glob.glob(workdir + '/' + company + '/' + resume_regexp)
        source[0] = source[0].replace('/', '\\')

        doc = self.winword.word.Documents.Open(source[0])
        doc.SaveAs(source[0].replace('docx', 'pdf').replace('tech', company), 17)
        doc.Close()

    def generate_cover_letter(self, workdir, company, position, job_portal):
        source = glob.glob(workdir + '/' + company + '/' + cover_letter_regexp)
        source[0] = source[0].replace('/', '\\')

        replacements = {
            '[Position Title]' : position,
            '[Company Name]' : company,
            '[Platform/Source]' : job_portal,
            '[Date]' : str(date.today())
        }

        doc = self.winword.word.Documents.Open(source[0])

        for find_text, replace_with in replacements.items():
            for paragraph in doc.Paragraphs:
                if find_text in paragraph.Range.Text:
                    paragraph.Range.HighlightColorIndex = 0
                    paragraph.Range.Text = paragraph.Range.Text.replace(find_text, replace_with)

        doc.SaveAs(source[0].replace('docx', 'pdf').replace('tech', company), 17)
        doc.Close()
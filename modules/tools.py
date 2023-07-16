import os
import shutil
import sys
import psutil
import glob
import win32com.client as win32
from config import email_regexp, resume_regexp, cover_letter_regexp


class Preparator:
    def __init__(self, workdir, company, job_type):
        self.workdir = workdir
        self.company = company
        self.job_type = job_type

    def prepare_dir(self):
        self._check_ifsent(self.workdir, self.company)
        self._makedir(self.workdir, self.company)
        self._copy_templates(self.workdir, self.company, self.job_type)

    def _check_ifsent(self, workdir, company):
        company_sent_dir = os.path.join(workdir, '_Sent', company)
        if os.path.isdir(company_sent_dir):
            ifcontinue = input('Вы уже отправляли резюме этой компании, продолжить? (y/n): ')
            if ifcontinue != 'y':
                sys.exit()

    def _makedir(self, workdir, company):
        company_dir = os.path.join(workdir, company)
        if not os.path.isdir(company_dir):
            os.makedirs(company_dir)

    def _copy_templates(self, workdir, company, job_type):
        templates_path = os.path.join(workdir + '/_templates')
        templates = glob.glob(templates_path + '/*.docx')
        templates += glob.glob(templates_path + '/*.txt')
        for template in templates:
            if job_type == 't' and template.find('tech') != -1:
                shutil.copy(template, workdir + '/' + company)
            elif job_type == 'm' and template.find('manager') != -1:
                shutil.copy(template, workdir + '/' + company)

class WinWordAdapter:
    def __init__(self):
        self.word = None

    def open_doc(self, source):
        self.doc = self.word.Documents.Open(source)

    def save_close_doc(self, ext_old, ext_new, type_res, company, pdf_code, source):
        self.doc.SaveAs(source.replace(ext_old, ext_new).replace(type_res, company), pdf_code)
        self.doc.Close()

    def open_word(self):
        self._terminate_word()
        self.word = win32.gencache.EnsureDispatch('Word.Application')
        self.word.Visible = False

    def close_word(self):
        self.word.Quit()

    def makepath(self, dir, regexp):
        source = glob.glob(dir + '/' + regexp)
        source[0] = source[0].replace('/', '\\')
        return source[0]

    def _terminate_word(self):
        for proc in psutil.process_iter():
            if proc.name() == 'WINWORD.EXE':
                input('Microsoft Word запущен, сохраните открытые документы. Enter для продолжения...')
                proc.terminate()

class GarbageRemover():
    def __init__(self, workdir, company):
        self.workdir = workdir
        self.company = company

    def final_clear(self):
        self._remove_docx(self.workdir, self.company)

    def _remove_docx(self, workdir, company):
        docxs = glob.glob(workdir + '/' + company + '/' + '*.docx')
        for docx in docxs:
            os.remove(docx)
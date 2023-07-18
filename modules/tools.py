''' it's a module for different tools '''
import os
import shutil
import sys
import glob

class Preparator:
    ''' this class is presented for folders prep and checking is resume\
          is already applied to this company '''
    def __init__(self, workdir, company, job_type):
        self.workdir = workdir
        self.company = company
        self.job_type = job_type

    def prepare_dir(self):
        ''' it's a main function to prepare dir '''
        self._check_ifsent()
        self._makedir()
        self._copy_templates()

    def _check_ifsent(self):
        company_sent_dir = os.path.join(self.workdir, '_Sent', self.company)
        if os.path.isdir(company_sent_dir):
            ifcontinue = input('Вы уже отправляли резюме этой компании, продолжить? (y/n): ')
            if ifcontinue != 'y':
                sys.exit()

    def _makedir(self):
        company_dir = os.path.join(self.workdir, self.company)
        if not os.path.isdir(company_dir):
            os.makedirs(company_dir)

    def _copy_templates(self):
        templates_path = os.path.join(self.workdir + '/_templates')
        templates = glob.glob(templates_path + '/*.docx')
        templates += glob.glob(templates_path + '/*.txt')
        for template in templates:
            if self.job_type == 't' and template.find('tech') != -1:
                shutil.copy(template, self.workdir + '/' + self.company)
            elif self.job_type == 'm' and template.find('manager') != -1:
                shutil.copy(template, self.workdir + '/' + self.company)


class GarbageRemover():
    ''' clear directory after job and leaving only needed files '''
    def __init__(self, workdir, company):
        self.workdir = workdir
        self.company = company

    def final_clear(self):
        ''' main func for call clearance '''
        self._remove_docx()

    def _remove_docx(self):
        docxs = glob.glob(self.workdir + '/' + self.company + '/' + '*.docx')
        for docx in docxs:
            os.remove(docx)

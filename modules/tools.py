''' it's a module for different tools '''
import os
import shutil
import sys
import glob
from config import resume_types

class WindowsTools:
    ''' it's a class for win tools '''
    def prep_path_for_win(self, path, regexp):
        ''' just preparing paths for windows '''
        source_path = glob.glob(path + '/' + regexp)
        source_path[0] = source_path[0].replace('/', '\\')
        return source_path[0]

class Preparator:
    ''' this class is presented for folders prep and checking is resume\
          is already applied to this company '''
    def __init__(self, companypath, workdir, company, job_type):
        self.workdir = workdir
        self.company = company
        self.companypath = companypath
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
        if not os.path.isdir(self.companypath):
            os.makedirs(self.companypath)

    def _copy_templates(self):
        templates_path = os.path.join(self.workdir + '/_templates')
        templates = glob.glob(templates_path + '/*.docx')
        templates += glob.glob(templates_path + '/*.txt')
        for template in templates:
            for key, value in resume_types.items():
                if self.job_type == key and template.find(value) != -1:
                    shutil.copy(template, self.companypath)


class GarbageRemover():
    ''' clear directory after job and leaving only needed files '''
    def __init__(self, path):
        self.path = path

    def final_clear(self):
        ''' main func for call clearance '''
        self._remove_docx()

    def _remove_docx(self):
        docxs = glob.glob(self.path + '/*.docx')
        for docx in docxs:
            os.remove(docx)

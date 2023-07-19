''' it's a usecases module '''
import os
import sys
from modules.tools import Preparator, GarbageRemover
from modules.generators import WinDocsGenerator

class WindowsWordCase:
    ''' it's a class used to UseCase for Windows and Word'''
    def __init__(self, workdir, company, job_type, position, job_portal):
        self.companypath = os.path.join(workdir, company)
        self.prepare = Preparator(self.companypath, workdir, company, job_type)
        self.clear = GarbageRemover(self.companypath)
        self.windocgen = WinDocsGenerator(self.companypath, company, job_type, position, job_portal)

    def make_documents(self):
        ''' it's a main function '''

        preconditions = self.prepare.check_preconditions()
        if preconditions['if_path_exists']:
            ifcontinue = input('Каталог клиента уже существует, продолжить работу? (y/n): ')
            if ifcontinue != 'y':
                sys.exit()
        if preconditions['if_already_sent']:
            ifcontinue = input('Вы уже отправляли резюме этой компании, продолжить? (y/n): ')
            if ifcontinue != 'y':
                sys.exit()
        if preconditions['word_processes'] == ['WINWORD.EXE']:
            ifcontinue = input('Microsoft Word запущен, в случае продолжения он будет закрыт. Продолжить? (y/n): ')
            if ifcontinue != 'y':
                sys.exit()

        self.prepare.prepare_dir()

        self.windocgen.generate()

        self.clear.final_clear()
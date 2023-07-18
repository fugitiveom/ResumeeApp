''' it's a usecases module '''
import os
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
        self.prepare.prepare_dir()

        self.windocgen.generate()

        self.clear.final_clear()

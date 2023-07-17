''' it's a usecases module '''
from modules.tools import Preparator, GarbageRemover
from modules.generators import WinDocsGenerator

class WindowsWordCase:
    ''' it's a class used to UseCase for Windows and Word'''
    def __init__(self, workdir, company, job_type, position, job_portal):
        self.prepare = Preparator(workdir, company, job_type)
        self.clear = GarbageRemover(workdir, company)
        self.windocgen = WinDocsGenerator(workdir, company, job_type, position, job_portal)

    def make_documents(self):
        ''' it's a main function '''
        self.prepare.prepare_dir()

        self.windocgen.generate()

        self.clear.final_clear()

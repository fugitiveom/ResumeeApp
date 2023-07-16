from modules.tools import Preparator, GarbageRemover, WinWord
from modules.generators import WinDocsGenerator

class WindowsWord_case:
    def __init__(self, workdir, company, job_type, position, job_portal):
        self.workdir = workdir
        self.company = company
        self.job_type = job_type
        self.position = position
        self.job_portal = job_portal
        self.prepare = Preparator(self.workdir, self.company, self.job_type)
        self.clear = GarbageRemover(self.workdir, self.company)
        self.winword = WinWord()
        self.windocgen = WinDocsGenerator(self.workdir, self.company, self.job_type, self.position, self.job_portal)

    def make_documents(self):
        self.prepare.prepare_dir()

        self.windocgen.generate()

        self.clear.final_clear()
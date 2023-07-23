''' it's a usecases module '''
import os
from dataclasses import dataclass
from modules.tools import Preparator, GarbageRemover, WindowsTools
from modules.generators import DocsGenerator
from modules.adapters import WinWordAdapter

@dataclass(frozen=True)
class UseCaseDataDTO:
    ''' DTO object for vars exchange '''
    def __init__(self, company, job_type, position, job_portal, replacements):
        object.__setattr__(self, 'company', company)
        object.__setattr__(self, 'job_type', job_type)
        object.__setattr__(self, 'position', position)
        object.__setattr__(self, 'job_portal', job_portal)
        object.__setattr__(self, 'replacements', replacements)

class WindowsWordCase:
    ''' it's a class used to UseCase for Windows and Word '''
    def __init__(self, workdir, data_dto: UseCaseDataDTO):
        self.data_dto = data_dto
        self.office_adapter=WinWordAdapter()
        self.tools = WindowsTools()
        self.companypath = os.path.join(workdir, self.data_dto.company)
        self.prepare = Preparator(self.companypath, workdir, self.data_dto.company, \
                                  self.data_dto.job_type)
        self.clear = GarbageRemover(self.companypath)
        self.docgen = DocsGenerator(self.companypath, self.data_dto, self.office_adapter, self.tools)

    def make_documents(self):
        ''' it's a main function '''
        try:
            self.prepare.prepare_dir()
        except OSError:
            print('При подготовке директории возникла ошибка. Выполняется откат изменений')
            self.clear.remove_directory()

        try:
            self.docgen.generate()
        except OSError:
            print('При генерации документов произошла ошибка. Выполняется откат изменений')
            self.clear.remove_directory()

        try:
            self.clear.final_clear()
        except OSError:
            print('При финальной очистке каталога произошла ошибка.')

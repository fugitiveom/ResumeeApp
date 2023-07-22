''' it's a usecases module '''
import os
import sys
from dataclasses import dataclass
from modules.tools import Preparator, GarbageRemover
from modules.generators import DocsGenerator

@dataclass(frozen=True)
class UseCaseDataDTO:
    ''' DTO object for vars exchange '''
    def __init__(self, company, job_type, position, job_portal, replacements):
        self.company = company
        self.job_type = job_type
        self.position = position
        self.job_portal = job_portal
        self.replacements = replacements

class WindowsWordCase:
    ''' it's a class used to UseCase for Windows and Word '''
    def __init__(self, workdir, data_dto: UseCaseDataDTO):
        self.data_dto = data_dto
        self.companypath = os.path.join(workdir, self.data_dto.company)
        self.prepare = Preparator(self.companypath, workdir, self.data_dto.company, \
                                  self.data_dto.job_type)
        self.clear = GarbageRemover(self.companypath)
        self.docgen = DocsGenerator(self.companypath, self.data_dto)

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
            ifcontinue = input('Microsoft Word запущен, в случае продолжения он будет закрыт. '
                               'Продолжить? (y/n): ')
            if ifcontinue != 'y':
                sys.exit()

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

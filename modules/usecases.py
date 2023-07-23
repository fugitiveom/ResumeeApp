''' it's a usecases module '''
import os
import sys
from modules.tools import Preparator, GarbageRemover, Tools
from modules.generators import DocsGenerator
from modules.adapters import WinWordAdapter
from modules.dto import UseCaseDataDTO

class WindowsWordCase:
    ''' it's a class used to UseCase for Windows and Word '''
    def __init__(self, workdir, data_dto: UseCaseDataDTO):
        self.data_dto = data_dto
        self.office_adapter=WinWordAdapter()
        self.tools = Tools()
        self.companypath = os.path.join(workdir, self.data_dto.company)
        self.prepare = Preparator(self.companypath, workdir, self.data_dto.company, \
                                  self.data_dto.job_type)
        self.clear = GarbageRemover(self.companypath)
        self.docgen = DocsGenerator(self.companypath, self.data_dto, self.office_adapter, self.tools)

    def make_documents(self):
        ''' it's a main function '''
        word_process = WinWordAdapter.if_word_open()

        if word_process == ['WINWORD.EXE']:
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

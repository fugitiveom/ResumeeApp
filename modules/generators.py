''' it's a generator's module'''
from config import EMAIL_REGEXP, RESUME_REGEXP, COVER_LETTER_REGEXP, resume_types, JOB_TYPE

class DocsGenerator():
    ''' this class is representing a generator of documents '''
    def __init__(self, gen_path, data_dto, office_adapter, tools):
        self.gen_path = gen_path
        self.data_dto = data_dto
        self.office_adapter = office_adapter
        self.tools = tools

    def generate(self):
        ''' main generating method '''
        self._generate_email_textfile()
        self._convert_resume_to_pdf()
        self._edit_cover_letter()

    def _generate_email_textfile(self):
        source_path = self.tools.prep_path_for_win(self.gen_path, EMAIL_REGEXP)
        with open(source_path, 'r+', encoding="UTF-8") as file:
            data = file.read()
            data = data.replace('[position name]', self.data_dto.position)
            data = data.replace('[Company Name]', self.data_dto.company)
            data = data.replace('[Job Source]', self.data_dto.job_portal)
            file.seek(0)
            file.write(data)
            file.truncate()

    def _convert_resume_to_pdf(self):
        type_res = resume_types[JOB_TYPE]
        source_path = self.tools.prep_path_for_win(self.gen_path, RESUME_REGEXP)
        new_path = source_path.replace(type_res, self.data_dto.company)
        self.office_adapter.save_docx_as_pdf(source_path, new_path)

    def _edit_cover_letter(self):
        type_res = resume_types[JOB_TYPE]
        source_path = self.tools.prep_path_for_win(self.gen_path, COVER_LETTER_REGEXP)

        self.office_adapter.replace_text_docx(source_path, self.data_dto.replacements)

        new_path = source_path.replace(type_res, self.data_dto.company)
        self.office_adapter.save_docx_as_pdf(source_path, new_path)

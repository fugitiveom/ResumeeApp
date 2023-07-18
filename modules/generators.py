''' it's a generator's module'''
from datetime import date
from config import EMAIL_REGEXP, RESUME_REGEXP, COVER_LETTER_REGEXP
from modules.adapters import WinWordAdapter
from modules.tools import WindowsTools

class WinDocsGenerator():
    ''' this class is representing a generator of documents '''
    def __init__(self, workdir, company, job_type, position, job_portal):
        self.workdir = workdir
        self.companydir = workdir + '/' + company
        self.company = company
        self.job_type = job_type
        self.position = position
        self.job_portal = job_portal
        self.winword = WinWordAdapter()
        self.wintools = WindowsTools()

    def generate(self):
        ''' main generating method '''
        self.winword.open_word()
        self._generate_email_textfile()
        self._convert_resume_to_pdf()
        self._edit_cover_letter()
        self.winword.close_word()

    def _generate_email_textfile(self):
        source_path = self.wintools.prep_path_for_win(self.companydir, EMAIL_REGEXP)
        with open(source_path, 'r+', encoding="UTF-8") as file:
            data = file.read()
            data = data.replace('[position name]', self.position)
            data = data.replace('[Company Name]', self.company)
            data = data.replace('[Job Source]', self.job_portal)
            file.seek(0)
            file.write(data)
            file.truncate()

    def _convert_resume_to_pdf(self):
        type_res = 'tech'
        source_path = self.wintools.prep_path_for_win(self.companydir, RESUME_REGEXP)
        self.winword.open_doc(source_path)
        new_path = source_path.replace(type_res, self.company)
        self.winword.save_docx_as_pdf(new_path)

    def _edit_cover_letter(self):
        type_res = 'tech'
        source_path = self.wintools.prep_path_for_win(self.companydir, COVER_LETTER_REGEXP)

        replacements = {
            '[Position Title]' : self.position,
            '[Company Name]' : self.company,
            '[Platform/Source]' : self.job_portal,
            '[Date]' : str(date.today())
        }

        self.winword.open_doc(source_path)

        for find_text, replace_with in replacements.items():
            for paragraph in self.winword.doc.Paragraphs:
                if find_text in paragraph.Range.Text:
                    paragraph.Range.HighlightColorIndex = 0
                    paragraph.Range.Text = paragraph.Range.Text.replace(find_text, replace_with)

        new_path = source_path.replace(type_res, self.company)
        self.winword.save_docx_as_pdf(new_path)

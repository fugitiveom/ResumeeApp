''' it's a generator's module'''
from datetime import date
from config import EMAIL_REGEXP, RESUME_REGEXP, COVER_LETTER_REGEXP
from modules.adapters import WinWordAdapter

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

    def generate(self):
        ''' main generating method '''
        self.winword.open_word()
        self._generate_email_textfile()
        self._convert_resume_to_pdf()
        self._edit_cover_letter()
        self.winword.close_word()

    def _generate_email_textfile(self):
        source = self.winword.makepath(self.companydir, EMAIL_REGEXP)
        with open(source, 'r+', encoding="UTF-8") as file:
            data = file.read()
            data = data.replace('[position name]', self.position)
            data = data.replace('[Company Name]', self.company)
            data = data.replace('[Job Source]', self.job_portal)
            file.seek(0)
            file.write(data)
            file.truncate()

    def _convert_resume_to_pdf(self):
        ext_old = 'docx'
        ext_new = 'pdf'
        type_res = 'tech'
        pdf_code = 17
        source = self.winword.makepath(self.companydir, RESUME_REGEXP)
        self.winword.open_doc(source)
        self.winword.save_close_doc(ext_old, ext_new, type_res, self.company, pdf_code, source)

    def _edit_cover_letter(self):
        ext_old = 'docx'
        ext_new = 'pdf'
        type_res = 'tech'
        pdf_code = 17
        source = self.winword.makepath(self.companydir, COVER_LETTER_REGEXP)

        replacements = {
            '[Position Title]' : self.position,
            '[Company Name]' : self.company,
            '[Platform/Source]' : self.job_portal,
            '[Date]' : str(date.today())
        }

        self.winword.open_doc(source)

        for find_text, replace_with in replacements.items():
            for paragraph in self.winword.doc.Paragraphs:
                if find_text in paragraph.Range.Text:
                    paragraph.Range.HighlightColorIndex = 0
                    paragraph.Range.Text = paragraph.Range.Text.replace(find_text, replace_with)

        self.winword.save_close_doc(ext_old, ext_new, type_res, self.company, pdf_code, source)

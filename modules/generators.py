import glob
from config import email_regexp, resume_regexp, cover_letter_regexp
from datetime import date
from modules.tools import WinWord

class WinDocsGenerator():
    def __init__(self, workdir, company, job_type, position, job_portal):
        self.workdir = workdir
        self.companydir = workdir + '/' + company
        self.company = company
        self.job_type = job_type
        self.position = position
        self.job_portal = job_portal
        self.winword = WinWord()

    def generate(self):
        self.winword.open_word()
        self._generate_email_textfile()
        self._convert_resume_to_pdf()
        self._edit_cover_letter()
        self.winword.close_word()

    def _generate_email_textfile(self):  # TODO Переписать на построчный вариант
        source = glob.glob(self.companydir + '/' + email_regexp)
        source[0] = source[0].replace('/', '\\')
        with open(source[0], 'r+') as f:
            data = f.read()
            data = data.replace('[position name]', self.position)
            data = data.replace('[Company Name]', self.company)
            data = data.replace('[Job Source]', self.job_portal)
            f.seek(0)
            f.write(data)
            f.truncate()

    def _convert_resume_to_pdf(self):
        source = glob.glob(self.companydir + '/' + resume_regexp)
        source[0] = source[0].replace('/', '\\')

        doc = self.winword.word.Documents.Open(source[0])
        doc.SaveAs(source[0].replace('docx', 'pdf').replace('tech', self.company), 17)
        doc.Close()

    def _edit_cover_letter(self):
        source = glob.glob(self.companydir + '/' + cover_letter_regexp)
        source[0] = source[0].replace('/', '\\')

        replacements = {
            '[Position Title]' : self.position,
            '[Company Name]' : self.company,
            '[Platform/Source]' : self.job_portal,
            '[Date]' : str(date.today())
        }

        doc = self.winword.word.Documents.Open(source[0])

        for find_text, replace_with in replacements.items():
            for paragraph in doc.Paragraphs:
                if find_text in paragraph.Range.Text:
                    paragraph.Range.HighlightColorIndex = 0
                    paragraph.Range.Text = paragraph.Range.Text.replace(find_text, replace_with)

        doc.SaveAs(source[0].replace('docx', 'pdf').replace('tech', self.company), 17)
        doc.Close()
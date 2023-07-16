import glob
from config import email_regexp, resume_regexp, cover_letter_regexp
from datetime import date
from modules.tools import WinWord

class WinDocsGenerator():
    def __init__(self, workdir, company, job_type, position, job_portal):
        self.workdir = workdir
        self.company = company
        self.job_type = job_type
        self.position = position
        self.job_portal = job_portal
        self.winword = WinWord()

    def generate(self):
        self.winword.open_word()
        self.generate_email(self.workdir, self.company, self.position, self.job_portal)
        self.generate_resume(self.workdir, self.company)
        self.generate_cover_letter(self.workdir, self.company, self.position, self.job_portal)
        self.winword.close_word()

    def generate_email(self, workdir, company, position, job_portal):  # TODO Переписать на построчный вариант
        source = glob.glob(workdir + '/' + company + '/' + email_regexp)
        source[0] = source[0].replace('/', '\\')
        with open(source[0], 'r+') as f:
            data = f.read()
            data = data.replace('[position name]', position)
            data = data.replace('[Company Name]', company)
            data = data.replace('[Job Source]', job_portal)
            f.seek(0)
            f.write(data)
            f.truncate()

    def generate_resume(self, workdir, company):
        source = glob.glob(workdir + '/' + company + '/' + resume_regexp)
        source[0] = source[0].replace('/', '\\')

        doc = self.winword.word.Documents.Open(source[0])
        doc.SaveAs(source[0].replace('docx', 'pdf').replace('tech', company), 17)
        doc.Close()

    def generate_cover_letter(self, workdir, company, position, job_portal):
        source = glob.glob(workdir + '/' + company + '/' + cover_letter_regexp)
        source[0] = source[0].replace('/', '\\')

        replacements = {
            '[Position Title]' : position,
            '[Company Name]' : company,
            '[Platform/Source]' : job_portal,
            '[Date]' : str(date.today())
        }

        doc = self.winword.word.Documents.Open(source[0])

        for find_text, replace_with in replacements.items():
            for paragraph in doc.Paragraphs:
                if find_text in paragraph.Range.Text:
                    paragraph.Range.HighlightColorIndex = 0
                    paragraph.Range.Text = paragraph.Range.Text.replace(find_text, replace_with)

        doc.SaveAs(source[0].replace('docx', 'pdf').replace('tech', company), 17)
        doc.Close()
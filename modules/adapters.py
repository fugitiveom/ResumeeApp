''' it's a module for adapters'''
import win32com.client as win32
import psutil

class WinWordAdapter:
    ''' this class is an adapter for WinWord '''
    def __init__(self):
        self.word = None
        self.pdf_type_no = 17
        self.word_proc = 'WINWORD.EXE'
        self._open_silent_word()

    @staticmethod
    def if_word_open():
        ''' checks if word process run '''
        word_process = [proc.name() for proc in psutil.process_iter() \
            if proc.name() == 'WINWORD.EXE']
        return word_process

    def save_docx_as_pdf(self, path, new_path=None):
        ''' save as PDF after files prepairing '''
        doc = self._open_doc(path)
        path = new_path if new_path else path
        doc.SaveAs(path.replace('docx', 'pdf'), self.pdf_type_no)
        doc.Close()

    def replace_text_docx(self, path, replacements):
        ''' replace text in doc '''
        doc = self._open_doc(path)
        for find_text, replace_with in replacements.items():
            for paragraph in doc.Paragraphs:
                if find_text in paragraph.Range.Text:
                    paragraph.Range.HighlightColorIndex = 0
                    paragraph.Range.Text = paragraph.Range.Text.replace(find_text, replace_with)
        doc.Save()
        doc.Close()

    def _open_doc(self, source_path):
        ''' open doc with WORD COM-obj '''
        doc = self.word.Documents.Open(source_path)
        return doc

    def _open_silent_word(self):
        ''' for speed optimizing we open Word globally at the start '''
        self._terminate_word()
        self.word = win32.gencache.EnsureDispatch('Word.Application')
        self.word.Visible = False

    def _close_word(self):
        ''' close Word after a job '''
        self.word.Quit()

    def _terminate_word(self):
        for proc in psutil.process_iter():
            if proc.name() == self.word_proc:
                proc.terminate()

    def __del__(self):
        self._close_word()

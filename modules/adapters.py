''' it's a module for adapters'''
import win32com.client as win32
import psutil

class WinWordAdapter:
    ''' this class is an adapter for WinWord '''
    def __init__(self):
        self.word = None
        self.doc = None
        self._open_word()

    def save_docx_as_pdf(self, path, new_path=None):
        ''' save as PDF after files prepairing '''
        self._open_doc(path)
        path = new_path if new_path else path
        self.doc.SaveAs(path.replace('docx', 'pdf'), 17)
        self.doc.Close()

    def replace_text_docx(self, path, replacements):
        ''' replace text in doc '''
        self._open_doc(path)
        for find_text, replace_with in replacements.items():
            for paragraph in self.doc.Paragraphs:
                if find_text in paragraph.Range.Text:
                    paragraph.Range.HighlightColorIndex = 0
                    paragraph.Range.Text = paragraph.Range.Text.replace(find_text, replace_with)
        self.doc.Save()
        self.doc.Close()

    def _open_doc(self, source_path):
        ''' open doc with WORD COM-obj '''
        self.doc = self.word.Documents.Open(source_path)

    def _open_word(self):
        ''' for speed optimizing we open Word globally at the start '''
        self._terminate_word()
        self.word = win32.gencache.EnsureDispatch('Word.Application')
        self.word.Visible = False

    def _close_word(self):
        ''' close Word after a job '''
        self.word.Quit()

    def _terminate_word(self):
        for proc in psutil.process_iter():
            if proc.name() == 'WINWORD.EXE':
                proc.terminate()

    def __del__(self):
        self._close_word()

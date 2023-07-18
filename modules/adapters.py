''' it's a module for adapters'''
import win32com.client as win32
import psutil

class WinWordAdapter:
    ''' this class is an adapter for WinWord '''
    def __init__(self):
        self.word = None
        self.doc = None

    def open_doc(self, source):
        ''' open doc with WORD COM-obj '''
        self.doc = self.word.Documents.Open(source)

    def save_docx_as_pdf(self, type_res, company, source):
        ''' save as PDF after files prepairing '''
        self.doc.SaveAs(source.replace('docx', 'pdf').replace(type_res, company), 17)
        self.doc.Close()

    def open_word(self):
        ''' for speed optimizing we open Word globally at the start '''
        self._terminate_word()
        self.word = win32.gencache.EnsureDispatch('Word.Application')
        self.word.Visible = False

    def close_word(self):
        ''' close Word after a job '''
        self.word.Quit()

    def _terminate_word(self):
        for proc in psutil.process_iter():
            if proc.name() == 'WINWORD.EXE':
                input('Microsoft Word запущен, сохраните открытые документы. Enter для продолжения...')
                proc.terminate()

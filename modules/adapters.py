''' it's a module for adapters'''
import glob
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

    def save_close_doc(self, ext_old, ext_new, type_res, company, pdf_code, source):
        ''' save as PDF after files prepairing '''
        self.doc.SaveAs(source.replace(ext_old, ext_new).replace(type_res, company), pdf_code)
        self.doc.Close()

    def open_word(self):
        ''' for speed optimizing we open Word globally at the start '''
        self._terminate_word()
        self.word = win32.gencache.EnsureDispatch('Word.Application')
        self.word.Visible = False

    def close_word(self):
        ''' close Word after a job '''
        self.word.Quit()

    def makepath(self, directory, regexp):
        ''' just preparing paths for windows '''
        source = glob.glob(directory + '/' + regexp)
        source[0] = source[0].replace('/', '\\')
        return source[0]

    def _terminate_word(self):
        for proc in psutil.process_iter():
            if proc.name() == 'WINWORD.EXE':
                input('Microsoft Word запущен, сохраните открытые документы. Enter для продолжения...')
                proc.terminate()

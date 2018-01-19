from win32com.client import gencache
from subprocess import Popen
from glob import glob

import win32com.client
import time
import os


EXPORT_DIR = 'C:\\Users\\GARY\\Desktop\\CAD\\'
INVENTOR_DIR = 'D:\\BC-Workspace\\GARY\\M-Balmoral,D-AGR\\Projects\\'
INVENTOR = 'C:\\Program Files\\Autodesk\\Inventor 2016\\Bin\\Inventor.exe'


def start_application():
    """Inventor Application Object
    Start COM client session with Inventor, and create object "mod" that will
    point to the Python COM wrapper for Inventor's type library. Recast "app"
    as an instance of the Application class in the wrapper.
    """
    if 'Inventor.exe' not in os.popen("tasklist").read():
        Popen(INVENTOR)
        time.sleep(5)
    mod = gencache.EnsureModule(
        '{D98A091D-3A0F-4C3E-B36E-61F62068D488}', 0, 1, 0)
    app = win32com.client.Dispatch('Inventor.Application')
    app = mod.Application.Application(app)
    app.SilentOperation = True
    app.Visible = True
    return app


class Document:

    def __init__(self, partcode):
        self.partcode = partcode
        self.filepath = self.inventor_filepath(partcode+'iam')
        self.doc = self.load(self.filepath)

    def inventor_filepath(self, filename, directory=INVENTOR_DIR):
        client = '*\\'
        project = filename[0:7] + '\\'
        section = filename[8:11] + '\\'
        paths = glob(directory + client + project + section + filename)
        if len(paths) > 1:
            print('Warning - multiple files found. Use first in the list')
        if len(paths) == 0:
            print('Unable to find ' + filename)
        if len(paths) > 0:
            return paths[0]

    def load(self, filename, app):
        """Inventor Document Object
        Check document type and bind the document COM object to the associate
        class in the wrapper.
        """
        doc_type = {
            12291: 'AssemblyDocument',
            12294: 'DesignElementDocument',
            12292: 'DrawingDocument',
            12295: 'ForeignModelDocument',
            12297: 'NoDocument',
            12290: 'PartDocument',
            12293: 'PresentationDocument',
            12296: 'SATFileDocument',
            12289: 'UnnownDocument',
        }
        app.Documents.Open(r'C:\Users\GARY\Desktop\AGR1197-105-00.idw')
        doc = win32com.client.CastTo(
            app.ActiveDocument, doc_type[app.ActiveDocumentType])
        print(doc_type[app.ActiveDocumentType])
        return doc

    def close(self):
        self.doc.Close()


def create_skeleton(partcode, directory=EXPORT_DIR):
    if not os.path.exists(directory+partcode):
        os.makedirs(directory + partcode + '\\pdf\\A0')


def main():
    start_application()


if __name__ == '__main__':
    main()

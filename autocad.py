from system import EXPORT_DIR
import win32com.client
import os


def application(visible=True):
    """Autcoad Application COM Object

    Start COM client session with Autocad.

    Parameters
    ----------
    silent : bool
        controls whether an operation will proceed without prompting
    visible : bool
        sets the visibility of this application

    Returns
    -------
    obj
        Autocad Application COM Object
    """

    app = win32com.client.Dispatch('AutoCAD.Application')
    app.Visible = visible
    return app


class Document:
    """Document

    The Document base class contains methods and properties for Autocad's
    Document COM object. Used to query inventor current file.
    Please Note! There is not much documentation online on AutoCAD COM API.
    End up using the AutoCAD 'SendCommand' method most of the time.
    Set the 'load file' prompt off so we can automate.

    Parameters
    ----------
    path : obj
        Path object from python pathlib module
    app : obj
        AutoCAD Application COM Object
    export_dir : str
        Export directory location

    Attributes
    ----------
    path : obj
        Path Object from python pathlib module
    app : obj
        AutoCAD Application COM Object
    doc : obj
        AutoCAD Document COM Object
    export_dir : str
        Export directory location
    """
    def __init__(self, path, app, export_dir=EXPORT_DIR):
        self.app = app
        self.doc = app.Documents.Open(str(path), False)
        self.doc.SendCommand('FILEDIA 0 ')
        self.export_dir = export_dir
        self.path = path


    @property
    def partcode(self):
        """str: return the file's partcode from it's full path"""
        return self.path.stem

    def export_to_pdf(self, subdir):
        """Export dwg file to pdf

        -EXPORT
        Enter file format [Dwf dwfX Pdf]
        Enter plot area [Display Extents Window]
        Detailed plot configuration [Yes No]
        Enter file name
        """
        file = self.partcode + '.' + 'pdf'
        path = self.export_dir.joinpath(subdir).joinpath(file)
        self.doc.SendCommand('-EXPORT PDF E NO {}\n'.format(str(path)))

    def export_to_dwf(self, subdir):
        """Export dwg file to dwg

        -Export
        Enter file format [Dwf dwfX Pdf]
        Enter plot area [Display Extents Window]
        Detailed plot configuration [Yes No]
        Enter file name
        """
        file = self.partcode + '.' + 'pdf'
        path = self.export_dir.joinpath(subdir).joinpath(file)
        self.doc.SendCommand('-EXPORT DWF E NO {}\n'.format(str(path)))

    def export_to_dxf(self, subdir):
        """Export dwg file to dxf

        DXFOUT
        Save drawing as <file location>
        Enter decimal places of accuracy (0 to 16)
        """
        file = self.partcode + '.' + 'dxf'
        path = self.export_dir.joinpath(subdir).joinpath(file)
        self.doc.SendCommand('DXFOUT {}\n16 '.format(str(path)))

    def export_to(self, subdir, filetype='pdf'):
        """Publish file into the export directory.

        Can export files such as; dwf, dxf, dwg and pdf.
        Use the following commands:

        -EXPORT
        Enter file format [Dwf dwfX Pdf]
        Enter plot area [Display Extents Window]
        Detailed plot configuration [Yes No]
        Enter file name

        DXFOUT
        Save drawing as <file location>
        Enter decimal places of accuracy (0 to 16)

        SAVE
        Save drawings as <file location>

        Parameters
        ----------
        subdir: str
            Sub directory for exported file.
        filetype: str
            Inventor supported file format
        """
        file = self.partcode + '.' + filetype
        path = self.export_dir.joinpath(subdir).joinpath(file)

        command = {
            'pdf': '-EXPORT PDF E NO {}\n',
            'dwf': '-EXPORT DWF E NO {}\n',
            'dxf': 'DXFOUT\n{}\n16 ',
            'dwg': 'SAVE\n{}\n'
        }
        self.doc.SendCommand(command[filetype].format(str(path)))

    def close(self):
        """Close Document

        Close current AutoCAD document without saving.
        Unable to automate the 'Do you want to save?' prompt. Around this,
        the file was temporary saved in anoother location and then deleted.
        """
        self.doc.SendCommand('SAVE C:\\delete_me.dwg\n')
        self.doc.SendCommand('FILEDIA 1 ')
        self.doc.SendCommand('CLOSE ')
        os.remove('C:\\delete_me.dwg')


class Drawing(Document):
    pass

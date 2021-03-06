"""
Inventor COM API
"""

from system import EXPORT_DIR, start_inventor
import win32com.client


class Document:
    """Document

    The Document base class contains methods and properties for Inventor's
    Document COM object. Used to query inventor current file.

    Parameters
    ----------
    path : obj
        Path object from python pathlib module
    app : obj
        Inventor Application COM Object
    export_dir : str
        Export directory location

    Attributes
    ----------
    path : obj
        Path Object from python pathlib module
    app : obj
        Inventor Application COM Object
    doc : obj
        Inventor Document COM Object
    export_dir : str
        Export directory location
    """

    def __init__(self, path, app, export_dir=EXPORT_DIR):
        self.app = app
        self.doc = self._load_document(path, app)
        self.export_dir = export_dir
        self.path = path
        print(self.doc)
        print(self.export_dir)
        print(str(self.path))

    @property
    def partcode(self):
        """str: return the file's partcode from it's full path"""
        return self.path.stem

    @staticmethod
    def _load_document(path, app):
        """Inventor Document Object

        Open the specified Inventor document.
        Check document type and bind the document COM object to the associate
        class in the wrapper.

        Parameters
        ----------
        path : obj
            Path Object from python pathlib module
        app : obj
            Inventor Application COM Object

        Returns
        -------
        obj
            Inventor Document COM Object
        """
        start_inventor()
        document_type_enum = {
            12289: 'UnnownDocument',
            12290: 'PartDocument',
            12291: 'AssemblyDocument',
            12292: 'DrawingDocument',
            12293: 'PresentationDocument',
            12294: 'DesignElementDocument',
            12295: 'ForeignModelDocument',
            12296: 'SATFileDocument',
            12297: 'NoDocument',
        }
        try:
            app.Documents.Open(str(path))
            document_type = document_type_enum[app.ActiveDocumentType]
            doc = win32com.client.CastTo(app.ActiveDocument, document_type)
            print(doc, document_type)
            return doc
        except:
            print('unable to load file')
            return None

    def export_to(self, subdir, filetype='pdf'):
        """Export file

        Publish file into the export directory using the Inventor translator
        add-in. Can export files such as; dwf, dxf, dwg, pdf, iges & step.

        Parameters
        ----------
        subdir: str
            Sub directory for exported file.
        filetype: str
            Inventor supported file format
        """
        file = self.partcode + '.' + filetype
        path = self.export_dir.joinpath(subdir).joinpath(file)
        print(str(path))
        self.doc.SaveAs(str(path), True)

    def close(self):
        """Close Document

        Close current Inventor document without saving.
        """
        self.doc.Close(SkipSave=True)


class Drawing(Document):
    """Drawing Document

    The Drawing Class contains methods and properties for Inventor's
    DrawingDocument COM object. Used to query idw file.
    """

    def get_drawing_sheet_size(self):
        """Sheet Size

        Get the size of the sheet.
        """
        drawing_sheet_size_enum = {
            9993: 'A0', 9994: 'A1', 9995: 'A2', 9996: 'A3', 9997: 'A4'
        }
        return drawing_sheet_size_enum[self.doc.Sheets(1).Size]

    def get_drawing_info(self):
        """Drawing Infomation

        Return a dictionary of the drawing's properties.
        """
        iprop = self.doc.PropertySets.Item("Inventor User Defined Properties")
        drawing_info = {
            'partcode': str(iprop.Item('Dwg_No')),
            'rev': int(iprop.Item('Revision')),
            'desc': str(iprop.Item('Component')),
            'material': str(iprop.Item('Material')),
            'finish': str(iprop.Item('Finish')),
            'size': self.get_drawing_sheet_size()
        }
        return drawing_info

    def export_part_list(self, filetype='xlsx'):
        """Part List

        Export the drawings part list to an excel spreadsheet.
        """
        if filetype == 'csv':
            enum = 48649
        else:
            enum = 48642
        path = self.export_dir.joinpath(self.partcode).joinpath('part_list.xlsx')
        self.doc.Sheets(1).PartsLists(1).Export(str(path), enum)


class Assembly(Document):
    """Assembly Document

    The Assembly Class contains methods and properties for Inventor's
    AssemblyDocument COM object. Used to query iam file.
    """

    def export_bom(self):
        """Assembly BOM

        Export the assembly's bom to an excel spreadsheet.
        """
        path = self.export_dir.joinpath(self.partcode).joinpath('bom.xlsx')
        bom = self.doc.ComponentDefinition.BOM
        bom.StructuredViewFirstLevelOnly = False
        bom.StructuredViewEnabled = True
        bom.BOMViews.Item("Structured").Export(path, 74498)


class Part(Document):
    """Part Document

    The Part Class contains methods and properties for Inventor's
    PartDocument COM object. Used to query ipt file.
    """
    pass


def application(silent=True, visible=True):
    """Inventor Application COM Object

    Start COM client session with Inventor, and create object 'mod' that will
    point to the Python COM wrapper for Inventor's type library. Recast 'app'
    as an instance of the Application class in the wrapper.

    Parameters
    ----------
    silent : bool
        controls whether an operation will proceed without prompting
    visible : bool
        sets the visibility of this application

    Returns
    -------
    obj
        Inventor Application COM Object
    """
    mod = win32com.client.gencache.EnsureModule(
        '{D98A091D-3A0F-4C3E-B36E-61F62068D488}', 0, 1, 0)
    app = win32com.client.Dispatch('Inventor.Application')
    app = mod.Application.Application(app)
    app.SilentOperation = silent
    app.Visible = visible
    return app

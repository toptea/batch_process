"""
Core Program
"""

import inventor
import autocad
import system

import pandas as pd
import zipfile
import os


def process_assembly(assembly, app):
    """Process Assembly

    Assembly must be an Inventor file and have pick list on the first page.

    1) Open assembly drawing (ipt)
    2) Pull drawing info to dict
    3) Save to spreadsheet - drawing_info.xlsx
    4) Export print, pdf and dxf files
    5) Save to spreadsheet - part_list.xlsx
    6) Open assembly part (iam)
    7) Save spreadsheet - bom.xlsx

    Parameters
    ----------
    assembly : str
        AGR part number usually in 'AGR0000-000-00' format
    app : obj
        Inventor Application COM Object
    """

    # 1) Open assembly drawing (ipt)
    path = system.find_path(assembly, 'idw')
    idw = inventor.Drawing(path, app)

    # 2) Pull drawing info to dict
    rs = pd.DataFrame(columns=['partcode', 'rev', 'desc', 'material', 'finish', 'size'])
    drawing_info = idw.get_drawing_info()
    rs = rs.append(drawing_info, ignore_index=True)

    # 3) Save to spreadsheet - drawing_info.xlsx
    path = system.EXPORT_DIR.joinpath(assembly).joinpath('drawing_info.xlsx')
    rs.to_excel(str(path), index=False)

    # 4) Export print, pdf and dxf file
    _export_inventer_drawing(idw, assembly, drawing_info, is_assy=True)

    # 5) Save to spreadsheet - part_list.xlsx
    idw.export_part_list('xlsx')
    idw.close()

    # 6) Open assembly part (iam)
    path = system.find_path(assembly, 'iam')
    iam = inventor.Assembly(path, app)

    # 7) Save spreadsheet - bom.xlsx
    iam.export_bom()
    iam.close()


def _export_inventer_drawing(idw, assembly, drawing_info, is_assy=False):
    """Export print, pdf and dxf files from one drawing

    Used in 'process_assembly(assembly, app)' and
    process_parts(assembly, app) functions.

    Parameters
    ----------
    idw: obj
        Drawing Document class from inventor.py
    assembly: str
        AGR part number usually in 'AGR0000-000-00' format
    drawing_info: dict
        Drawing information
    is_assy: bool
        is the part an assenmbly?
    """
    print_size = drawing_info['size']
    if print_size == 'A2':
        print_size = 'A3'
    print_dir = assembly + '/print/' + print_size + '/'
    idw.export_to(print_dir, 'pdf')
    idw.export_to(assembly + '/pdf/', 'pdf')
    if not is_assy:
        # idw.export_to(assembly + '/dxf/', 'dxf')
        # idw.export_to(assembly + '/dwg/', 'dwg')
        pass

def _load_children(assembly):
    """Load children from parent

    Used in 'create_format_matrix(assembly)' function.
    From the assembly idw's part list and iam's bom, return a list
    of all the drawings used under this section.

    Parameters
    ----------
    assembly: str
        AGR part number usually in 'AGR0000-000-00' format

    Returns
    -------
    partcodes: 'obj' of 'str'
        list of drawings
    """
    path1 = system.EXPORT_DIR.joinpath(assembly).joinpath('part_list.xlsx')
    df1 = pd.read_excel(str(path1))
    p1 = df1.loc[df1['Dwg_No'].notnull(), 'Dwg_No']

    path2 = system.EXPORT_DIR.joinpath(assembly).joinpath('bom.xlsx')
    df2 = pd.read_excel(str(path2))
    p2 = df2.loc[df2['Part Number'].notnull(), 'Part Number']

    partcodes = [*p1, *p2]
    partcodes = list(dict.fromkeys(partcodes))
    return partcodes


def create_format_matrix(assembly):
    """File format type spreadsheeet

    Find all ipt, iam, idw and dwg files under the assembly,
    create a file format report.

    Parameters
    ----------
    assembly: str
        AGR part number usually in 'AGR0000-000-00' format
    """
    ipt_paths = []
    iam_paths = []
    idw_paths = []
    dwg_paths = []
    partcodes = _load_children(assembly)
    for partcode in partcodes:
        ipt_paths.append(system.find_path(partcode, 'ipt') is not None)
        iam_paths.append(system.find_path(partcode, 'iam') is not None)
        idw_paths.append(system.find_path(partcode, 'idw') is not None)
        dwg_paths.append(system.find_path(partcode, 'dwg') is not None)

    df = pd.DataFrame()
    df['partcode'] = partcodes
    df['ipt'] = ipt_paths
    df['iam'] = iam_paths
    df['idw'] = idw_paths
    df['dwg'] = dwg_paths

    path = system.EXPORT_DIR.joinpath(assembly).joinpath('format_type.xlsx')
    df.to_excel(str(path), index=False)


def process_parts(assembly, app):
    """Process Parts

    1) Load spreadsheet - drawing_info.xlsx
    2) Load spreadsheet - format_type.xlsx
    3) Create a list of idw paths
    4) Open each drawings (idw)
    5) Pull drawing info to dict
    6) Export print, pdf and dxf files
    7) Save spreadsheet - drawing_info.xlsx
    8) Create a list of dwg paths
    9) export pdf files (AutoCAD)
    10) Unzip files

    Parameters
    ----------
    assembly : str
        AGR part number usually in 'AGR0000-000-00' format
    app : obj
        Inventor Application COM Object
    """
    # 1) Load spreadsheet - drawing_info.xlsx
    info_path = system.EXPORT_DIR.joinpath(assembly).joinpath('drawing_info.xlsx')
    if os.path.exists(str(info_path)):
        rs = pd.read_excel(str(info_path))
    else:
        rs = pd.DataFrame(columns=['partcode', 'rev', 'desc', 'material', 'finish', 'size'])

    # 2) Load spreadsheet -  format_type.xlsx, create a list of idw paths
    path = system.EXPORT_DIR.joinpath(assembly).joinpath('format_type.xlsx')
    df = pd.read_excel(str(path))
    inv_df = df.loc[df['idw']==True, ['partcode', 'iam']]
    atc_df = df.loc[df['dwg']==True, ['partcode']]

    # 3) Create a list of idw paths
    paths = []
    for partcode in inv_df['partcode']:
        path = system.find_path(partcode, 'idw')
        paths.append(path)

    # 4) Open each drawings
    # 5) Pull drawing info to dict
    # 6) export print, pdf and dxf files
    # 7) Save spreadsheet - drawing_info.xlsx
    if len(paths) > 0:
        for path, is_assy in zip(paths, df['iam']):
            idw = inventor.Drawing(path, app)
            drawing_info = idw.get_drawing_info()
            rs = rs.append(drawing_info, ignore_index=True)
            _export_inventer_drawing(idw, assembly, drawing_info, is_assy)
            idw.close()

        rs.to_excel(str(info_path), index=False)

    # 8) Create a list of dwg paths
    paths = []
    for partcode in atc_df['partcode']:
        path = system.find_path(partcode, 'dwg')
        paths.append(path)

    # 9) export pdf files (AutoCAD)
    autocad_app = autocad.application()
    if len(paths) > 0:
        for path in paths:
            dwg = autocad.Drawing(path, autocad_app)
            dwg.export_to(assembly + r'\from_autocad', 'pdf')
            dwg.close()

    # 10) Unzip files
    # for partcode in inv_df['partcode']:
    #     file = partcode + '.zip'
    #     export_path = system.EXPORT_DIR.joinpath(assembly).joinpath('dxf')
    #     with zipfile.ZipFile(str(export_path.joinpath(file)), 'r') as zip_ref:
    #         zip_ref.extractall(str(export_path))
    #     os.remove(str(export_path.joinpath(file)))


def batch_export(assembly):
    """Main - Batch Export Drawing

    1) Create project folder
    2) Connect to Inventor COM API
    3) Process assembly
    4) Find all idw part files in the vault
    5) Process parts
    """
    system.create_project(assembly)
    app = inventor.application()
    process_assembly(assembly, app)
    create_format_matrix(assembly)
    process_parts(assembly, app)


def batch_export_from(filename, filetype):
    app = inventor.application()

    with open(str(system.EXPORT_DIR.joinpath(filename))) as file:
        partcodes = [line.strip() for line in file]

    paths = []
    ipt_convert = [
        'CATPart', 'jt', 'ipt', 'igs', 'iges', 'sat',
        'smt', 'stl', 'step', 'stp', 'xgl', 'zgl'
    ]
    for partcode in partcodes:
        if filetype in ipt_convert:
            path = system.find_path(partcode, 'ipt')
        else:
            path = system.find_path(partcode, 'idw')
        paths.append(path)

    for path in paths:
        if filetype in ipt_convert:
            inv = inventor.Part(path, app)
        else:
            inv = inventor.Drawing(path, app)
        inv.export_to(system.EXPORT_DIR, filetype)
        inv.close()

    for partcode in partcodes:
        try:
            file = partcode + '.zip'
            export_path = system.EXPORT_DIR
            with zipfile.ZipFile(str(export_path.joinpath(file)), 'r') as zip_ref:
                zip_ref.extractall(str(export_path))
            os.remove(str(export_path.joinpath(file)))
        except:
            pass


def export_to(partcode, filetype):
    """Main - Export to ...

    Export one drawing to the specified file format
    """
    app = inventor.application()
    ipt_convert = [
        'CATPart', 'jt', 'ipt', 'igs', 'iges', 'sat',
        'smt', 'stl', 'step', 'stp', 'xgl', 'zgl'
    ]
    if filetype in ipt_convert:
        path = system.find_path(partcode, 'ipt')
        inv = inventor.Part(path, app)
    else:
        path = system.find_path(partcode, 'idw')
        inv = inventor.Drawing(path, app)
    inv.export_to(system.EXPORT_DIR, filetype)
    inv.close()

    try:
        file = partcode + '.zip'
        export_path = system.EXPORT_DIR
        with zipfile.ZipFile(str(export_path.joinpath(file)), 'r') as zip_ref:
            zip_ref.extractall(str(export_path))
        os.remove(str(export_path.joinpath(file)))
    except:
        pass


def cnc_batch_export_from(filename):
    """export dxf files and save them on a memory stick for cnc"""
    batch_export_from(filename, 'dxf')
    with open(str(system.EXPORT_DIR.joinpath(filename))) as file:
        partcodes = [line.strip() for line in file]

    for partcode in partcodes:
        file = partcode + '.dxf'
        path = system.EXPORT_DIR.joinpath(file)
        os.remove("G:/" + file)
        os.rename(str(path), "G:/" + file)


def cnc_export_to(partcode):
    """export dxf file and save them on a memory stick for cnc"""
    export_to(partcode, 'dxf')
    file = partcode + '.dxf'
    path = system.EXPORT_DIR.joinpath(file)
    os.remove("G:/" + file)
    os.rename(str(path), "G:/" + file)

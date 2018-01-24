import pandas as pd
import system
import inventor
import os
import zipfile


def export_assembly(assembly, app):
    """export part list and assembly drawing"""
    path = system.find_path(assembly, 'idw')
    idw = inventor.Drawing(path, app)

    rs = pd.DataFrame(columns=['partcode', 'rev', 'desc', 'material', 'finish', 'size'])
    drawing_info = idw.get_drawing_info()
    rs = rs.append(drawing_info, ignore_index=True)
    path = system.EXPORT_DIR.joinpath(assembly).joinpath('drawing_info.xlsx')
    rs.to_excel(str(path), index=False)
    _export_drawing(idw, assembly, drawing_info, is_assy=True)
    idw.export_part_list('xlsx')
    idw.close()

    path = system.find_path(assembly, 'iam')
    iam = inventor.Assembly(path, app)
    iam.export_bom()
    iam.close()


def _export_drawing(idw, assembly, drawing_info, is_assy=False):
    print_size = drawing_info['size']
    if print_size == 'A2':
        print_size = 'A3'
    print_dir = assembly + '/print/' + print_size + '/'
    idw.export_to(print_dir, 'pdf')
    idw.export_to(assembly + '/pdf/', 'pdf')
    if not is_assy:
        idw.export_to(assembly + '/dxf/', 'dxf')
        # idw.export_to(assembly + '/dwg/', 'dwg')


def _load_children(assembly):
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


def export_parts(assembly, app):
    info_path = system.EXPORT_DIR.joinpath(assembly).joinpath('drawing_info.xlsx')
    if os.path.exists(str(info_path)):
        rs = pd.read_excel(str(info_path))
    else:
        rs = pd.DataFrame(columns=['partcode', 'rev', 'desc', 'material', 'finish', 'size'])

    path = system.EXPORT_DIR.joinpath(assembly).joinpath('format_type.xlsx')
    df = pd.read_excel(str(path))
    df = df.loc[df['idw']==True, ['partcode', 'iam']]

    paths = []
    for partcode in df['partcode']:
        path = system.find_path(partcode, 'idw')
        paths.append(path)

    for path, is_assy in zip(paths, df['iam']):
        idw = inventor.Drawing(path, app)
        drawing_info = idw.get_drawing_info()
        rs = rs.append(drawing_info, ignore_index=True)
        _export_drawing(idw, assembly, drawing_info, is_assy)
        idw.close()

    rs.to_excel(str(info_path), index=False)

    for partcode in df['partcode']:
        file = partcode + '.zip'
        export_path = system.EXPORT_DIR.joinpath(assembly).joinpath('dxf')
        with zipfile.ZipFile(str(export_path.joinpath(file)), 'r') as zip_ref:
            zip_ref.extractall(str(export_path))
        os.remove(str(export_path.joinpath(file)))


def batch_export(assembly):
    system.create_project(assembly)
    app = inventor.application()
    export_assembly(assembly, app)
    create_format_matrix(assembly)
    export_parts(assembly, app)


def export_to(partcode, filetype):
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
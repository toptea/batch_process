import pandas as pd
import system
import inventor
from pprint import pprint


def export_part_list(partcode, app):
    path = system.find_path(partcode, 'idw')
    idw = inventor.Drawing(path, app)
    idw.export_part_list('xlsx')
    idw.close()


def load_children(partcode):
    path = system.EXPORT_DIR.joinpath(partcode).joinpath('part_list.xlsx')
    df = pd.read_excel(str(path))
    partcodes = df.loc[df['Dwg_No'].notnull(), 'Dwg_No'].unique()
    return partcodes


def create_format_matrix(assembly):
    ipt_paths = []
    iam_paths = []
    idw_paths = []
    dwg_paths = []
    partcodes = load_children(assembly)
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
    df.to_excel(str(path))


def export_drawing_info(assembly, app):
    rs = pd.DataFrame(columns=['partcode', 'rev', 'desc', 'material', 'finish', 'size'])
    path = system.EXPORT_DIR.joinpath(assembly).joinpath('format_type.xlsx')
    df = pd.read_excel(str(path))
    partcodes = df.loc[df['idw']==True, 'partcode']

    paths = []
    for partcode in partcodes:
        path = system.find_path(partcode, 'idw')
        paths.append(path)

    for path in paths:
        idw = inventor.Drawing(path, app)
        row = idw.get_drawing_info()
        rs = rs.append(row, ignore_index=True)
        print(row)
        idw.close()

    path = system.EXPORT_DIR.joinpath(assembly).joinpath('drawing_info.xlsx')
    rs.to_excel(str(path))


def batch_export(partcode):
    system.create_project(partcode)
    app = inventor.application()
    export_part_list(partcode, app)
    create_format_matrix(partcode)
    export_drawing_info(partcode, app)


if __name__ == '__main__':
    batch_export('AGR1316-112-00')

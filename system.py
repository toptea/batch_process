"""
Operating System Methods
"""

from pathlib import Path
from glob import glob

import subprocess
import time
import os


EXPORT_DIR = Path('C:/Users/GARY/Desktop/CAD/')
INVENTOR_DIR = Path('D:/BC-Workspace/GARY/M-Balmoral,D-AGR/Projects')
INVENTOR_APP = Path('C:/Program Files/Autodesk/Inventor 2016/Bin/Inventor.exe')


def create_project(partcode):
    """Create project skeleton

    Parameters
    ----------
    partcode : str
        AGR part number usually in 'AGR0000-000-00' format
    """
    directory = str(EXPORT_DIR.joinpath(partcode))
    if not os.path.exists(directory):
        os.makedirs(directory)
        os.makedirs(directory + r'\print\A0')
        os.makedirs(directory + r'\print\A1')
        os.makedirs(directory + r'\print\A3')
        os.makedirs(directory + r'\pdf')
        os.makedirs(directory + r'\from_autocad')
        # os.makedirs(directory + r'\dxf')
        # os.makedirs(directory + '\dwg')


def find_path(partcode, filetype):
    """find Inventor file path

    Parameters
    ----------
    partcode : str
        AGR part number usually in 'AGR0000-000-00' format
    filetype : str
        inventor file type ('ipt', 'iam' or 'idw')

    Returns
    -------
    Path : obj
        Path object from Python pathlib module
    """
    client = '*'
    project = partcode[0:7]
    section = partcode[8:11]
    file = partcode + '.' + filetype
    paths = glob(str(INVENTOR_DIR / client / project / section / file))

    # if len(paths) == 0:
    #     paths = glob(str(INVENTOR_DIR / '**' / file), recursive=True)
    if len(paths) == 0:
        print('Unable to find ' + partcode)
    if len(paths) > 1:
        print('Warning - multiple files found. Use first path in list')
    if len(paths) > 0:
        return Path(paths[0])


def find_paths(partcodes, filetype):
    """find Inventor file paths

    Parameters
    ----------
    partcodes : :obj:`list` of :obj:`str`
        AGR part numbers usually in 'AGR0000-000-00' format
    filetype : str
        inventor file type ('ipt', 'iam' or 'idw')

    Returns
    -------
    Path : :obj:`list` of obj
        Path object from Python pathlib module
    """
    paths = []
    for partcode in partcodes:
        path = find_path(partcode, filetype)
        if path is not None:
            paths.append(path)
    return paths


def start_inventor():
    """Open Inventor

    Inventer must be active for the COM API to work.
    """
    if 'Inventor.exe' not in os.popen("tasklist").read():
        subprocess.Popen(INVENTOR_APP)
        time.sleep(10)


if __name__ == '__main__':
    pass
    # start_inventor()
    # create_project('Hello')
    # print(find_paths(['AGR1316-010-00', 'AGR1316-020-00'], 'iam'))

import win32com.client
from pprint import pprint
from win32com.client import gencache
from glob import glob
import os

EXPORT_DIR = 'C:\\Users\\GARY\\Desktop\\CAD\\'
INVENTOR_DIR = 'D:\\BC-Workspace\\GARY\\M-Balmoral,D-AGR\\Projects\\'


def inventor_filepath(filename, directory=INVENTOR_DIR):
    client = '*\\'
    project = filename[0:7] + '\\'
    section = filename[8:11] + '\\'
    paths = glob(directory+client+project+section+filename)
    if len(paths) > 1:
        print('Warning - multiple files found. Use first in the list')
    if len(paths) == 0:
        print('Unable to find'+filename)
    if len(paths) > 0:
        return paths[0]


def create_skeleton(partcode, directory=EXPORT_DIR):
    if not os.path.exists(directory+partcode):
        os.makedirs(directory + partcode + '\\pdf\\A0')
        os.makedirs(directory + partcode + '\\pdf\\A1')
        os.makedirs(directory + partcode + '\\pdf\\A2')
        os.makedirs(directory + partcode + '\\pdf\\A3')
        os.makedirs(directory + partcode + '\\xlsx')




def main():
    create_skeleton('AGR1316-012-00')


if __name__ == '__main__':
    main()

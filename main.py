"""
User Interface
"""

import core


def title():
    print(
        """
        -------------------------------------------------------------
        ______       _       _       _____                      _
        | ___ \     | |     | |     |  ___|                    | |
        | |_/ / __ _| |_ ___| |__   | |____  ___ __   ___  _ __| |_
        | ___ \/ _` | __/ __| '_ \  |  __\ \/ / '_ \ / _ \| '__| __|
        | |_/ / (_| | || (__| | | | | |___>  <| |_) | (_) | |  | |_
        \____/ \__,_|\__\___|_| |_| \____/_/\_\ .__/ \___/|_|   \__|
                                              | |
                                              |_|
        -------------------------------------------------------------
                                by Gary Ip
        """
    )


def main_ui():
    print(
        """
        Welcome!

        This program provide the user an easy way to publish drawings
        from Autodesk Inventor. Just type in the assembly number for
        batch export, or the partcode for the exporting one file only.

        Before you start, please make sure:
        a) Save and close any active inventor file you have opened.
        b) All Inventor files are synchronized in Meridian.
        c) The BOM/part list on the assembly's iam/idw file can be accessed.
        d) Input/Output paths are correct in setup.ini

        Please pick the following options:
        1) Batch Export
        2) Batch Export from <file> to <file format>
        3) Export <partcode> to <file format>
        4) CNC - Batch Export dxf from <file> into G: Drive
        5) CNC - Export <partcode> to dxf into G: Drive
        q) quit
        """
    )

    while True:
        user_input = input('Export: ')
        if user_input == '1':
            batch_export_ui()
        elif user_input == '2':
            batch_export_from_ui()
        elif user_input == '3':
            export_to_ui()
        elif user_input == '4':
            cnc_batch_export_ui()
        elif user_input == '5':
            cnc_export_to()
        elif user_input == 'q' or user_input == 'quit':
            exit()
        else:
            print('Please enter a valid option')


def batch_export_ui():
    print("    Please enter an assembly number you want to export.")
    print("    Enter 'b' to go back")
    print("    Enter 'q' to quit\n")
    while True:
        user_input = input('Batch export: ')
        if user_input == 'b' or user_input == 'back':
            main_ui()
        elif user_input == 'q' or user_input == 'quit':
            exit()
        else:
            core.batch_export(user_input)


def batch_export_from_ui():
    print("    Please enter the file name with the list of partcodes you want to export.")
    print("    Enter 'b' to go back")
    print("    Enter 'q' to quit\n")
    while True:
        filename = input('File: ')
        if filename == '':
            filename = 'export.txt'
        if filename == 'b' or filename == 'back':
            main_ui()
        elif filename == 'q' or filename == 'quit':
            exit()
        else:
            print("    Please enter the file format you want to convert to.")
            print("    Enter 'b' to go back")
            print("    Enter 'q' to quit\n")
            filetype = input('File format: ')
            if filetype == 'b' or filetype == 'back':
                main_ui()
            elif filetype == 'q' or filetype == 'quit':
                exit()
            else:
                core.batch_export_from(filename, filetype)


def export_to_ui():
    print("    Please enter the drawing number you want to export.")
    print("    Enter 'b' to go back")
    print("    Enter 'q' to quit\n")
    while True:
        partcode = input('Export: ')
        if partcode == 'b' or partcode == 'back':
            main_ui()
        elif partcode == 'q' or partcode == 'quit':
            exit()
        else:
            print("    Please enter the file format you want to convert to.")
            print("    Enter 'b' to go back")
            print("    Enter 'q' to quit\n")
            filetype = input('File format: ')
            if filetype == 'b' or filetype == 'back':
                main_ui()
            elif filetype == 'q' or filetype == 'quit':
                exit()
            else:
                core.export_to(partcode, filetype)


def cnc_batch_export_ui():
    print("    Please enter the file name with the list of partcodes you want to export.")
    print("    Enter 'b' to go back")
    print("    Enter 'q' to quit\n")
    while True:
        filename = input('File: ')
        if filename == '':
            filename = 'export.txt'
        if filename == 'b' or filename == 'back':
            main_ui()
        elif filename == 'q' or filename == 'quit':
            exit()
        else:
            core.cnc_batch_export_from(filename)


def cnc_export_to():
    print("    Please enter the drawing number you want to export.")
    print("    Enter 'b' to go back")
    print("    Enter 'q' to quit\n")
    while True:
        partcode = input('Export: ')
        if partcode == 'b' or partcode == 'back':
            main_ui()
        elif partcode == 'q' or partcode == 'quit':
            exit()
        else:
            core.cnc_export_to(partcode)


if __name__ == '__main__':
    title()
    main_ui()

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
        2) Export to...
        q) quit
        """
    )

    while True:
        user_input = input('Export: ')
        if user_input == '1':
            batch_export_ui()
        elif user_input == '2':
            export_to_ui()
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
            filetype = input('File format: ')
            if filetype == 'b' or filetype == 'back':
                main_ui()
            elif filetype == 'q' or filetype == 'quit':
                exit()
            else:
                core.export_to(partcode, filetype)


if __name__ == '__main__':
    title()
    main_ui()


import os
from colorama import Fore, Style
from colorama import init as colorama_init

class UserQuitException(Exception):
    pass

def reverse_dict(dict):
    new_dict = {}

    for key,val in dict.items():
        new_dict[ val ] = key

    return new_dict

def make_menu(options):
    """Makes a command-line menu, and the user only gets
    away when he selects a valid option.

    :param options: A list/tuple of his options, like ['Leave', 'Forward'].
    :return: the name of the option selected, like 'Forward'. """
    print('---------------------')
    for i, op in enumerate(options):
        print( Fore.YELLOW + f'[{i:^3}]:' + Style.RESET_ALL + f' {op}')
    print('---------------------')
    n_op = len(options)
    while True:
        try:
            selected = int(input('>>>\t'))

            if selected >= 0 and selected < n_op:
                return options[selected]

        except ValueError:
            print('Choose one of the above options, only with numbers!')

def scan_directory(extensions, path=os.getcwd()):
    """Returns a list of files in a directory."""
    final_list = []

    for filename in os.listdir(path):
        for extension in extensions:
            if str(filename).upper().endswith(extension.upper()):
                final_list.append(filename)
                break

    return final_list

def choose_file(extensions, path=os.getcwd()):
    """ Asks the user to choose between a file  """
    option = 'Rescan'

    while True:
        file_list = scan_directory( extensions, path)
        menu_options = ['Leave', 'Rescan', 'Change Path'] + file_list

        print(f'Current Path is "{path}"')
        option = make_menu(menu_options)

        if   option == 'Leave':
            raise UserQuitException()
        elif option == 'Change Path':
            path = input('Type in new path\n>>>\t')
        elif option in file_list:
            return os.path.join(path,option)

def autowidth_excel_columns(dataframe, worksheet):
    """ Sets the width of each column as 2 characters more than the largest string. """
    columnNames = dataframe.columns.values.tolist()

    for i,col in enumerate(dataframe.columns):
        max_len = max([len(str(item)) for item in dataframe[col].values])
        max_len = max([max_len, len(str(columnNames[i]))])
        # (i+1) because the first column is used for row indexes.
        worksheet.set_column(i+1, i+1, width=max_len+2)

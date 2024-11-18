''' DiffCount. Compares two reports of billing/wordcount and makes a comparison report. '''

import logging
import pandas
import numpy
import diffcount.util as util
import diffcount.columns as columns
from diffcount.util import UserQuitException
from datetime import datetime
from colorama import Fore, Style
from colorama import init as colorama_init
from sys import stdout

from xlsxwriter.utility import xl_col_to_name

def add_logging_handler(logging_filename, logging_level = logging.DEBUG):
    '''With the logging module already configured, appends a file as logging output.'''
    logging_file = logging.FileHandler(logging_filename)
    logging_file.setLevel(logging_level)
    logging_file.setFormatter(logging.Formatter(fmt='[%(asctime)s %(levelname)s] %(message)s',datefmt='%H:%M:%S'))

    logging.log(logging.INFO, f'Setting up logging file as [{logging_filename}]')

    logger = logging.getLogger()
    logger.addHandler(logging_file)

def get_dataframe(filename, column_mapping):
    '''Loads a XLSX file as a dataframe, pulling the columns in the mapping,
        and renaming them as the mapping asks.

        Example:
        column_mapping = { 'value' : 'Item Value'}
        will take the column named 'Item Value' and make it accessible as
        df['value'].
    '''
    logging.log(logging.INFO, f"Reading dataframe from file: {filename}")
    loaded_file = pandas.read_excel(filename)

    original_cols = column_mapping.values()
    for col in original_cols:
        if col not in loaded_file.columns:
            print( Fore.YELLOW + f"Error! " + Style.RESET_ALL + f"  Invalid Excel Format. Column named <{col}> not found.")
            # The line below will throw a convenient exception. This error message above is just more user-friendly.

    data = loaded_file[ (column_mapping.values()) ].copy()
    data.rename(columns=util.reverse_dict(column_mapping), inplace=True)

    return data

def fix_sol_target(row):
    '''For a target like 'enUS', we need to convert it to 'en_US'.'''
    # These do not receive an underscore.
    exceptions = ['ckb', 'ceb', 'eo', 'hil', 'cnr']

    if row not in exceptions:
        row = row[:2] + "_" + row[2:]

    if row == 'cnr':
        row = 'sla_ME'

    return row

def fix_tew_target(row):
    '''Change the name of the target so that it matches the SOL one.'''
    renaming_values = {
        'sr_RS_Latn' : 'sr_RS',
        'az_AZ_Latn' : 'az_AZ',
        'bs_BA_Latn' : 'bs_BA',
        'sr_RS_Cyrl' : 'sr_CP'
    }

    if row in list(renaming_values.keys()):
        row = renaming_values[row]

    return row

def fix_sol_dataframe(data):
    # There are a lot of zeroed rows (~5x the content rows) that really need to be filtered
    # But apparently, there may be relevant negative values.
    data = data[data.sol_total != 0].copy()
    data.loc[:,'tgroup']= 'SAPLSP' + data.tgroup
    data.target = data.target.apply(lambda x: fix_sol_target(x))

    # There are some rows in SOL without any LSP, while in TEW that isn't the case.
    # They are used only for internal recordkeeping, and therefore not for billing.
    data = data[ data['tgroup'].isnull() == False ]

    # Latinization does not affect billing, so it is irrelevant to us.
    data = data[ data['step'] != 'LATINIZE']
    del data['step']

    # Every line will have a SLS_UNIT, and we need to have each SLS_UNIT count in a different
    # column. It follows, then, that for each line, we copy the billing quantity to the matching
    # SLS_UNIT-specific column.
    for new_name, unit_key_name in columns.SOL_UNIT_NAME_MAPPING.items():
        # If a row has SLS_UNIT = 100_MATCH, then the column '100_MATCH' is assigned as billing quantity.
        data[new_name] = \
            numpy.select(
               [data.sls_unit != unit_key_name, True],
               [0, data.sol_total])

    # Sometimes the billing quantity is not summed in words, but in hours.
    # We need to discard those. Also, sol_total_unit is not needed afterwards.
    data.loc[data['sol_total_unit'] == 'H', 'sol_total'] = 0
    del data['sol_total_unit']

    # After the above change, we do not need the SLS_UNIT column anymore,
    # and it needs to be removed before the GROUP BY
    del data['sls_unit']

    data = data.groupby(columns.SOL_GROUPBY_COLUMNS, as_index=False).agg( 'sum' )

    return data

def fix_tew_dataframe(data):
    # There are a lot of zeroed rows (~5x the content rows) that really need to be filtered
    # But apparently, there may be relevant negative values.
    data = data[data.tew_total != 0].copy()
    data.target = data.target.apply(lambda x: fix_tew_target(x))

    data = data.groupby(columns.TEW_GROUPBY_COLUMNS, as_index=False).agg( 'sum' )

    # We need to sum the fuzzy match and the fuzzy repeats, each pair at the time
    # STATS_VOLUME_MEDIUM_FUZZY_MATCH_WORDS + STATS_VOLUME_MEDIUM_FUZZY_REPEATS_WORDS, for example.
    # In the dataframe, they are named tew_85_a, tew_85_b, respectively.
    # So we will create a column tew_85, which is their sum, and delete the two original columns.
    to_sum_cols = ['tew_95', 'tew_85', 'tew_75']
    for col in to_sum_cols:
        name_a = col + '_a'
        name_b = col + '_b'

        data[col] = data[name_a] + data[name_b]

        del data[name_a]
        del data[name_b]
        pass

    return data

def style_and_save(data, filename):

    #===================================
    # 1. Reorder all columns
    #===================================

    cols = data.columns.values.tolist()
    try:
        for col_in_order in columns.FINAL_COLUMN_ORDER:
            cols.remove(col_in_order)
        # '+cols' to make so that columns outside the
        # constant still exist, but in the end.
        cols = columns.FINAL_COLUMN_ORDER + cols
    except ValueError as ex:
        print('Trying to get columns that do not exist in the final col order. ')
        print(set(columns.FINAL_COLUMN_ORDER) - set(cols))

    data = data[cols]

    # AESTHETIC CHOICE
    show_as_indexes = False
    # It is basically a matter of presentation whether we
    # want to set a column with a simple index, if the join
    # columns should themselves be the index.
    if show_as_indexes:
        data.set_index(columns.JOIN_KEY_COLUMNS, inplace=True)

    #===================================
    # 2. Create an Excel Writer
    #===================================

    writer = pandas.ExcelWriter(filename, engine='xlsxwriter')
    data.to_excel(writer)

    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']
    max_row = data.shape[0]

    cell_format = workbook.add_format()
    cell_format.set_bold()
    cell_format.set_font_color('red')

    bad_format = workbook.add_format({'bg_color': '#FFC7CE',
                            'font_color': '#9C0006'})

    ok_format = workbook.add_format({'bg_color': '#C6EFCE',
                            'font_color': '#006100'})

    weird1_format = workbook.add_format({'bg_color': '#CCC0DA',
                            'font_color': '#403151'})

    weird2_format = workbook.add_format({'bg_color': '#FDE9D9',
                            'font_color': '#974706'})

    # weird3_format = workbook.add_format({'bg_color': '#FFA0FF',
    #                        'font_color': '#800080'})

    # Annoying column displaying an index specific to
    # the pandas datasheet.
    if not show_as_indexes:
        worksheet.set_column('A:A', options={'hidden':True})

    #===================================
    # 3. Conditional Formatting of 'Ok' and 'Mismatch' strings.
    #===================================
    # Rules have a 'last added' priority.

    worksheet.conditional_format(1, 5, max_row, 5,
                                {'type':     'cell',
                                    'criteria': 'equal to',
                                    'value':     '"MISMATCH"',
                                    'format':    bad_format})

    worksheet.conditional_format(1, 5, max_row, 5,
                            {'type':     'cell',
                                'criteria': 'equal to',
                                'value':     '"OK"',
                                'format':    ok_format})

    worksheet.conditional_format(1, 5, max_row, 5,
                            {'type':     'cell',
                                'criteria': 'equal to',
                                'value':     '"TEW_MISSING"',
                                'format':    weird1_format})

    worksheet.conditional_format(1, 5, max_row, 5,
                            {'type':     'cell',
                                'criteria': 'equal to',
                                'value':     '"SOL_MISSING"',
                                'format':    weird2_format})

    util.autowidth_excel_columns(data,worksheet)

    #===================================
    # 4. Conditional Formatting of Number Cols.
    #===================================

    for sol_col_title, tew_col_title in columns.FINAL_COLUMN_PAIRS:
        sol_col_i = columns.FINAL_COLUMN_ORDER.index(sol_col_title)
        tew_col_i = columns.FINAL_COLUMN_ORDER.index(tew_col_title)

        # Adding one because the first column in Excel is the index.
        if not show_as_indexes:
            sol_col_i += 1
            tew_col_i += 1

        sol_col = xl_col_to_name(sol_col_i)
        tew_col = xl_col_to_name(tew_col_i)
        worksheet.conditional_format(1,tew_col_i, max_row, tew_col_i, {
                        'type':     'formula',
                        'criteria':f'=${sol_col}2<>${tew_col}2',
                        'format':   bad_format
                    })

    # Makes the first row always visible, even if the user scrolls down.
    worksheet.freeze_panes(1, 0)

    writer.close()
    logging.log(logging.INFO, f'Saved file successfully as {filename}.')

def fix_merged_dataframe(data):

    # We by default take the name of the sol report, but sometimes
    # it is missing, so we fix that by using the TEW name instead.
    data.loc[data['projectName'].isnull(), 'projectName'] = data.loc[data['projectName'].isnull(), 'tew_name']
    del data['tew_name']

    # We create a status column based on if totals match, or if one of
    # the reports does not have this data point.
    data['status'] = numpy.select([
        (data.sol_total.isnull()),
        (data.tew_total.isnull()),
        (data.tew_total == data.sol_total),
        (data.tew_total != data.sol_total)],
        ["SOL_MISSING", "TEW_MISSING","OK", "MISMATCH"])

    # After merging, we do not want to left blank cells instead of zeroes.
    data.loc[data['tew_total'].isnull(), 'tew_total'] = 0
    data.loc[data['sol_total'].isnull(), 'sol_total'] = 0


    data.sort_values(by=['projectId'], inplace=True)

    return data

def sanity_check_data(data,tew_data,sol_data):
    """ Convenience printing after we have already done everything.
    This function could also be used to check 'data' against the two original files, too.  """
    print("\nRows by status:")
    status_types = ["SOL_MISSING", "TEW_MISSING", "OK", "MISMATCH"]
    for type in status_types:
        print(f"\t{type:20} : { len( data.loc[ data.status == type ] ) }")

    print(f"\n\t{'Total':20} : { len(data) } \n")
    pass

def main():
    colorama_init()
    # Greeting
    print(Fore.YELLOW + 'DiffCount.' + Style.RESET_ALL)
    print('SAP LX tool for comparing billing requests by SOL and TEW.')
    print('Hosted at https://github.wdf.sap.corp/I567680/DiffCount.')
    print('--------------------------------------------------\n\n')

    logging.basicConfig(level=logging.DEBUG, style='{', datefmt='%H:%M:%S', format='[{asctime} {levelname}] {message}', stream=stdout)
    add_logging_handler(f'diffcount_logs.log') # log-diffcount-{datetime.utcnow().strftime("%Y-%m-%d")}.log

    #===================================
    # 1. SOL File
    #===================================
    while True:
        try:
            print( Fore.YELLOW + 'Step 1.' + Style.RESET_ALL + ' Please select the SOL file [.xls, .xlsx].')
            sol_filename = util.choose_file( ('.xls', '.xlsx', '.csv') )
            sol_data = get_dataframe(sol_filename, columns.SOL_COLUMN_MAPPING)
            sol_data = fix_sol_dataframe(sol_data)

            break
        except PermissionError:
            print('Oh no! Permission denied!\nClose the file in excel so we can proceed.')
            answer = util.make_menu(['Try Again', 'Quit'])
            if answer == 'Quit':
                return
        except KeyError as ex:
            print(f'Oh no! It seems you have gotten the wrong file! {ex}')
            answer = util.make_menu(['Try Again', 'Quit'])
            if answer == 'Quit':
                return
        except UserQuitException as ex:
            print(f'\nGoodbye!')
            return


    logging.log(logging.INFO, f'Success! There are {sol_data.shape[0]} rows in the processed SOL file.')

    #===================================
    # 2. TEW File
    #===================================
    while True:
        try:
            print( Fore.YELLOW + 'Step 2.' + Style.RESET_ALL + ' Please select the TEW file [.xls, .xlsx].')
            tew_filename = util.choose_file( ('.xls', '.xlsx', '.csv') )
            tew_data = get_dataframe(tew_filename, columns.TEW_COLUMN_MAPPING)
            tew_data = fix_tew_dataframe(tew_data)
            break
        except PermissionError:
            print('Oh no! Permission denied!\nClose the file in excel so we can proceed.')
            answer = util.make_menu(['Try Again', 'Quit'])
            if answer == 'Quit':
                return
        except KeyError as ex:
            print(f'Oh no! Maybe it\'s the pivot table, but it seems you have gotten the wrong file! {ex}')
            answer = util.make_menu(['Try Again', 'Quit'])
            if answer == 'Quit':
                return
        except UserQuitException as ex:
            print(f'\nGoodbye!')
            return

    logging.log(logging.INFO, f'Success! There are {tew_data.shape[0]} rows in the processed TEW file.')

    #===================================
    # 3. Merging Data
    #===================================
    result_filename = 'diffcount-' + datetime.utcnow().strftime('%Y-%m-%d') + '.xlsx'
    print( Fore.YELLOW + 'Step 3.' + Style.RESET_ALL + f' Creating report file [{result_filename}].')

    data = pandas.merge(sol_data, tew_data, 'outer', on=columns.JOIN_KEY_COLUMNS)
    data = fix_merged_dataframe(data)

    logging.log(logging.INFO, f'Success! The resulting file {result_filename} will have {data.shape[0]} rows.')

    # Save the file.
    try:
        style_and_save(data, result_filename)
    except PermissionError:
        print('Oh no! Permission denied!\nClose the file in excel so we can proceed.')
        answer = util.make_menu(['Try Again', 'Quit'])
        if answer == 'Quit':
            return
    except Exception as ex:
        logging.exception(f'Could not save resulting file! {ex} ')
        return

    #===================================
    # 3. Checking and Printing
    #===================================
    sanity_check_data(data,tew_data,sol_data)

if __name__ == '__main__':
    main()
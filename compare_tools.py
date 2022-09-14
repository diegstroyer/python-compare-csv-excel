from pathlib import Path
import chardet
import pandas as pd

# * Autodetect file encoding
def file_encoding(filename):    
    rawdata = open(filename, 'rb').readline()
    result = chardet.detect(rawdata)
    charenc = result['encoding']
    return charenc

# * Read Excel files
def read_excel_to_df(excel_doc, index_col):
    # excel_doc = Path(excel_doc)
    excel_df = pd.read_excel(excel_doc, index_col=index_col).fillna('')
    # Delete all Unnamed columns
    excel_df = excel_df.loc[:, ~excel_df.columns.str.contains('^Unnamed:')]
    return excel_df

# * Read CSV files
def read_csv_to_df(csv_doc, index_col, delimiter, encoding):
    # TODO: Sometimes, file_encoding function not works, investigate it...
    try:
        # Try to use file_encoding function
        encoding_auto=file_encoding(csv_doc)
        csv_df = pd.read_csv(csv_doc,
                            encoding=encoding_auto,
                            delimiter=delimiter,
                            dtype=str,
                            index_col=index_col).fillna('')
    except:
        csv_df = pd.read_csv(csv_doc,
                            encoding=encoding,
                            delimiter=delimiter,
                            dtype=str,
                            index_col=index_col).fillna('')

    # Delete all Unnamed columns
    csv_df = csv_df.loc[:, ~csv_df.columns.str.contains('^Unnamed:')]
    return csv_df

# * Fromat EXCEL
def row_format_excel(row_list, df_checked, worksheet, dimensions, format):
    for value in row_list:
        rowvalue = df_checked.index.get_loc(value)
        worksheet.conditional_format((rowvalue + 1),
                                    0,(rowvalue + 1),
                                    dimensions[1],
                                    {'type': 'formula',
                                        'criteria': '=$A${}="{}"'.format(rowvalue + 2, value),
                                        'format': format})

# * Compute files and check differences
def diff_files_check(df_NEW, df_OLD):
    # Perform Diff
    dfDiff = df_NEW.copy()
    droppedRows = []
    newRows = []
    diffRows = []
    cols_OLD = df_OLD.columns
    cols_NEW = df_NEW.columns
    sharedCols = list(set(cols_OLD).intersection(cols_NEW))

    # Add status ROW (added, dropped or changed)
    dfDiff['STATUS'] = pd.NaT

    for row in dfDiff.index:
        if (row in df_OLD.index):
            for col in sharedCols:
                value_OLD = df_OLD.loc[row, col]
                value_NEW = df_NEW.loc[row, col]
                if value_OLD == value_NEW:
                    dfDiff.loc[row, col] = df_NEW.loc[row, col]
                else:
                    dfDiff.loc[row, col] = ('{}→{}').format(value_OLD, value_NEW)
                    dfDiff.loc[row, 'STATUS'] = 'CANVI' # Changed
                    diffRows.append(row)
        else:
            newRows.append(row)
            dfDiff.loc[row, 'STATUS'] = 'NOVA' # Added

    for row in df_OLD.index:
        if row not in df_NEW.index:
            droppedRows.append(row)
            dfDiff = dfDiff.append(df_OLD.loc[row, :])
            dfDiff.loc[row, 'STATUS'] = 'ELIMINADA' # Dropped

    # Print new and dropped rows
    dfDiff = dfDiff.sort_index().fillna('')
    print('\nNew Rows {}:   {}'.format(len(newRows), sorted(newRows)))
    print('\nDropped Rows {}:   {}'.format(len(droppedRows), sorted(droppedRows)))
    return dfDiff, newRows, droppedRows

# * Analize differences in files, final result
# * Excel
def excel_diff(path_OLD, path_NEW, index_col):
    path_OLD = Path(path_OLD)
    path_NEW = Path(path_NEW)
    df_OLD = read_excel_to_df(path_OLD, index_col)
    df_NEW = read_excel_to_df(path_NEW, index_col)

    # Use function diff_files_check to perform the changes test
    dfDiff, newRows, droppedRows = diff_files_check(df_NEW, df_OLD)

    # Save output and format
    fname = '{}_vs_{}.xlsx'.format(path_OLD.stem, path_NEW.stem)
    writer = pd.ExcelWriter(fname, engine='xlsxwriter')

    # Create worksheets with data
    dfDiff.to_excel(writer, sheet_name='DIFF', index=True)
    df_NEW.to_excel(writer, sheet_name=path_NEW.stem, index=True)
    df_OLD.to_excel(writer, sheet_name=path_OLD.stem, index=True)

    # Get xlsxwriter objects
    workbook = writer.book
    worksheet = writer.sheets['DIFF']

    # Define formats
    # date_fmt = workbook.add_format({'align': 'center', 'num_format': 'yyyy-mm-dd'})
    # center_fmt = workbook.add_format({'align': 'center'})
    # number_fmt = workbook.add_format({'align': 'center', 'num_format': '#,##0.00'})
    # cur_fmt = workbook.add_format({'align': 'center', 'num_format': '$#,##0.00'})
    # perc_fmt = workbook.add_format({'align': 'center', 'num_format': '0%'})
    # grey_fmt = workbook.add_format({'font_color': '#E0E0E0'})
    drop_fmt = workbook.add_format({'bold': True, 'font_color': '#ffffff', 'bg_color': '#FF0000'})
    highlight_fmt = workbook.add_format({'bold': True, 'font_color': '#FF0000', 'bg_color': '#B1B3B3'})
    new_fmt = workbook.add_format({'bold': True, 'font_color': '#ffffff', 'bg_color': '#00ae00'})

    # Extract dimension of DIFF worksheet, for apply format to data cells
    dimensions = dfDiff.shape

    # Apply format based in data cells (conditional format)
    # Changed values
    worksheet.conditional_format(0,
                                0,
                                dimensions[0],
                                dimensions[1],
                                {'type': 'text',
                                    'criteria': 'containing',
                                    'value': '→',
                                    'format': highlight_fmt})

    # New rows
    row_format_excel(newRows, dfDiff, worksheet, dimensions, new_fmt)
    
    # Dropped rows
    row_format_excel(droppedRows, dfDiff, worksheet, dimensions, drop_fmt)

    # Save
    writer.save()
    print('\nDone.\n')

# * CSV
def csv_diff(path_OLD, path_NEW, index_col, delimiter=';', encoding='ISO-8859-1'):
    path_OLD = Path(path_OLD)
    path_NEW = Path(path_NEW)

    df_OLD = read_csv_to_df(path_OLD, index_col, delimiter, encoding)
    df_NEW = read_csv_to_df(path_NEW, index_col, delimiter, encoding)

    # Use function diff_files_check to perform the changes test
    dfDiff = diff_files_check(df_NEW, df_OLD)[0]
    fname = '{}_vs_{}.csv'.format(path_OLD.stem, path_NEW.stem)
    dfDiff.to_csv(fname, index=True, sep= delimiter)
    print('\nDone.\n')

# Uncomment and change paths and index_col for test, be carefull with encoding option, default value is ISO-8859-1 and delimiter semicolon
# def main():
#     path_OLD = '20210705.xlsx'
#     path_NEW = '20220425.xlsx'
#     index_col = 'NIF'

#     excel_diff(path_OLD, path_NEW, index_col)
#     csv_diff(path_OLD, path_NEW, index_col)


# if __name__ == '__main__':
#     main()




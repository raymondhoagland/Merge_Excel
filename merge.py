import xlrd
import xlwt
import os
import sys
import re
import csv

default_output_columns = ["Company", "Week", "G/L#", "Acct Description", "Dept#", "Dept Description", "Debit", "Credit", "Description"]
allowed_filetypes = ["xlsx", "xls", "csv"]
preferred_output_extension = "xls"
tmp_filename = "merge_tmp"
output_sheet_name = "All Details Data"

# Add a single file to the set to merge
def include_file(filepath, fileset):
    global allowed_filetypes

    try:
        path = ''.join(re.findall(".*/", filepath))
        filename_no_path = re.sub(".*/", "", filepath)
        extension = filename_no_path.split('.')[-1]
    except IndexError as e:
        raise IOError("Please specify all file extensions.")

    # ignore if extension not supported
    if not (extension.lower() in allowed_filetypes):
        return fileset

    filename_to_add = filepath

    # if csv convert to xls and save as temp file
    if extension.lower() == "csv":
        output = "./tmpMERGE"+filename_no_path.split('.')[0]+".xls"
        print "Found csv: "+filepath
        print "Creating temporary file: "+output
        with open(filepath, 'rb') as csvfile:
            dest_workbook = xlwt.Workbook()
            dest_sheet = dest_workbook.add_sheet("Sheet1")
            reader = csv.reader(csvfile)
            for r, row in enumerate(reader):
                for c, col in enumerate(row):
                    val = col
                    # try converting to a float to remove warnings
                    try:
                        val = float(''.join(col.split(',')))
                    except:
                        pass
                    dest_sheet.write(r, c, val)
        dest_workbook.save(output)
        filename_to_add = output
    try:
        fileset.add(filename_to_add)
        return fileset
    except IOError as e:
        raise IOError("File not found.")

# add a full folder to the set to merge from
def include_folder(path, fileset):
    global allowed_filetypes

    # list out each file in the provided path and add those with supported extensions
    try:
        for filename in os.listdir(path):
            try:
                extension = filename.split('.')[-1]
            except IndexError as e:
                continue
            if not (extension in allowed_filetypes):
                continue
            fileset = include_file(path+filename, fileset)
        return fileset
    except IOError as e:
        raise IOError("Folder not found.")

# loop through input file to discover all files that needed to be merged
def retrieve_filenames(input_filepath):
    with open(input_filepath, 'rb') as input_file:
        included_files = set()
        modulo = 0
        op = -1
        for line in input_file:
            # os.listdir provides newlines which are unwanted
            normalized_line = line.rstrip('\n')
            lower_line = normalized_line.lower()
            # switch off in pairs of type-item
            if modulo == 0:
                modulo = 1
                if "file" in lower_line:
                    op = 0
                elif "dir" in lower_line or "folder" in lower_line:
                    op = 1
                else:
                    op = -1
            else:
                modulo = 0
                if op == 0:
                    # include next file
                    included_files = include_file(normalized_line, included_files)
                elif op == 1:
                    # include whole dir
                    included_files = include_folder(normalized_line, included_files)
                else:
                    # maybe throw error
                    continue
        return included_files

# retrieve the week and company from the filename
def split_filename(filename):
    filename_no_path = re.sub(".*/", "", filename)
    filename_no_ext = ''.join(filename_no_path.split('.')[:-1])
    print "Adding {0} to output.".format(filename_no_path)
    return tuple(filename_no_path.split(' '))

# check which sheet to write merged data to
def get_worksheet_index(src_workbook, key=output_sheet_name):
    for sheet_idx in xrange(src_workbook.nsheets):
        sheet = src_workbook.sheet_by_index(sheet_idx)
        if sheet.name == key:
            return sheet_idx
    raise IOError("Sheet not found in workbook")

# clone existing workbook that will be merged to (for subsequent runs)
def copy_existing(src):
    dest_workbook = xlwt.Workbook()
    src_workbook = xlrd.open_workbook(src)
    n_worksheets = src_workbook.nsheets
    merge_worksheet = None
    merge_col = -1
    merge_row = -1
    merge_idx = get_worksheet_index(src_workbook)
    headers = []
    # iterate through sheets
    for sheet_idx in xrange(n_worksheets):
        src_worksheet = src_workbook.sheet_by_index(sheet_idx)
        n_cols = src_worksheet.ncols
        n_rows = src_worksheet.nrows
        src_sheet_name = src_worksheet.name
        dest_ws = dest_workbook.add_sheet(src_sheet_name)
        # iterate through columns, rows
        for row in xrange(n_rows):
            for col in xrange(n_cols):
                cell_val = src_worksheet.cell_value(row, col)
                dest_ws.write(row, col, cell_val)
                if row == 0:
                    headers.append(cell_val)
        # if this is the sheet we want to merge into, identify the last used line
        if sheet_idx == merge_idx:
            merge_worksheet = dest_ws
            merge_col = n_cols
            merge_row = n_rows
    return (dest_workbook, merge_worksheet, merge_row, merge_col, headers)

# start off a new workbook by injecting default column headers
def init_output():
    global output_sheet_name, default_output_columns

    dest_workbook = xlwt.Workbook()
    dest_ws = dest_workbook.add_sheet(output_sheet_name)
    for col in xrange(len(default_output_columns)):
        dest_ws.write(0, col, default_output_columns[col])
    return (dest_workbook, dest_ws, 1, len(default_output_columns), default_output_columns)

# extract headers for existing workbooks
def collect_headers(src_worksheet):
    headers = []
    for col in xrange(src_worksheet.ncols):
        headers.append(src_worksheet.cell_value(0, col))
    return headers

# check in which order to write columns from the individual workbooks to the merged output
def map_headers(columns_src, columns_dest):
    mapping = [-1 for i in xrange(len(columns_src))]
    for col in xrange(len(columns_src)):
        try:
            mapping[col] = columns_dest.index(columns_src[col].rstrip())
        except ValueError as e:
            continue
    return mapping

# copy individual files into the output
def merge_files(dest_workbook, dest_ws, dest_headers, fileset, start_row, num_columns):
    global preferred_output_extension
    global tmp_filename

    dest_row, dest_col = start_row, 0
    for filename in fileset:
        (week, company) = split_filename(filename)
        with xlrd.open_workbook(filename, 'rb') as src_workbook:
            src_worksheet = src_workbook.sheet_by_index(0)
            src_headers = collect_headers(src_worksheet)
            mapping = map_headers(src_headers, dest_headers)
            for row in xrange(1, src_worksheet.nrows):
                dest_ws.write(dest_row, 0, company)
                dest_ws.write(dest_row, 1, week)
                for col in xrange(src_worksheet.ncols):
                    cell_val = src_worksheet.cell_value(row, col)
                    if mapping[col] != -1:
                        dest_ws.write(dest_row, mapping[col], cell_val)
                dest_row += 1
    # save to temporary file so we can delete an exisitng file if present
    dest_workbook.save(tmp_filename+preferred_output_extension)

if __name__ == '__main__':
    if len(sys.argv) < 2:
        raise IOError("Invalid args, please specify the output file path.")
    output_file = sys.argv[1]
    # validate output file extension
    try:
        path = ''.join(re.findall(".*/", output_file))
        filename_no_path = re.sub(".*/", "", output_file)
        _extension = filename_no_path.split('.')
        if len(_extension) > 1:
            if not (_extension[1] == preferred_output_extension):
                print "Warning: The output file will be stored as a ."+preferred_output_extension+ " file."
        else:
            output_file += "."+preferred_output_extension
    except IndexError as e:
        print 'Using default extension type of '+preferred_output_extension
    fileset = retrieve_filenames('./listing.txt')
    need_copy = os.path.isfile(output_file)
    # if the workbook exists, copy the contents, otherwise initialize
    if need_copy:
        (dest_workbook, dest_ws, merge_row, merge_col, dest_headers) = copy_existing(output_file)
    else:
        (dest_workbook, dest_ws, merge_row, merge_col, dest_headers) = init_output()
    merge_files(dest_workbook, dest_ws, dest_headers, fileset, merge_row, merge_col)
    # overwrite if needed
    if need_copy:
        os.remove(output_file)
    os.rename(tmp_filename+preferred_output_extension, output_file)
    # remove all the temp source files
    tmp_files_regex = re.compile("tmpMERGE.*")
    for f in os.listdir("./"):
        if tmp_files_regex.search(f):
            print "Removing temporary file {0}".format(f)
            os.remove("./"+f)

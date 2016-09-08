# Pre-requisites
python 2.7

xlrd

xlwt

# If you have python 2.7 and pip but not xlrd and xlwt
`pip install -r requirements.txt`

# Modifying which files are merged
Within listing.txt add 1+ pairs of lines using these two formats
1)
```
FILE
path_to_file_to_merge
```

2)
```
DIR
path_to_folder_to_merge
```

File types allowed are .csv, .xls, and .xlsx
If dir is specified, all files with .csv, .xls, or .xlsx extensions in the directory will be merged
All files explicitly specified must include the extension and a path (use ./ to represent the current path)
All directories must include the path (use ./ to represent the current path)

# Running the command
`python merge.py path_to_output_file`

The output file will always be stored as .xls

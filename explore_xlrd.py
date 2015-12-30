import xlrd

path = 'C:\\Users\\tech5\\Desktop\\PHOENIX RISING workout\\NEPR PHOENIX_RISING (2L2015) Long Beach-Tokyo.xlsx'

# load the book into python
book = xlrd.open_workbook(path)

# get descriptive info
ship_name = book.sheet_by_index(2).cell(0, 5).value
voyage_number = book.sheet_by_index(2).cell(1, 5).value
load_condition = book.sheet_by_index(2).cell(2, 5).value
cargo_desc = book.sheet_by_index(2).cell(3, 5).value
cargo_weight = book.sheet_by_index(2).cell(4, 5).value

departure_port = book.sheet_by_index(2).cell(0, 10).value
departure_date = xlrd.xldate_as_tuple(book.sheet_by_index(2).cell(1, 10).value, book.datemode)
departure_draft_fore = book.sheet_by_index(2).cell(2, 10).value
departure_draft_mid = book.sheet_by_index(2).cell(3, 10).value
departure_draft_aft = book.sheet_by_index(2).cell(4, 10).value

arrival_port = book.sheet_by_index(2).cell(0, 14).value
arrival_date = xlrd.xldate_as_tuple(book.sheet_by_index(2).cell(1, 14).value, book.datemode)
arrival_draft_fore = book.sheet_by_index(2).cell(2, 14).value
arrival_draft_mid = book.sheet_by_index(2).cell(3, 14).value
arrival_draft_aft = book.sheet_by_index(2).cell(4, 14).value

# get arrays
date_array = [xlrd.xldate_as_tuple(e, book.datemode) for e in book.sheet_by_index(2).col_values(colx=0, start_rowx=8) if e != '']
nrows = len(date_array)
time_array = [xlrd.xldate_as_tuple(e, book.datemode) for e in book.sheet_by_index(2).col_values(colx=1, start_rowx=8, end_rowx=8 + nrows)]
utc_time_array = [xlrd.xldate_as_tuple(e, book.datemode) for e in book.sheet_by_index(2).col_values(colx=2, start_rowx=8, end_rowx=8 + nrows)]
lat_array = [e for e in book.sheet_by_index(2).col_values(colx=3, start_rowx=8, end_rowx=8 + nrows)]
long_array = [e for e in book.sheet_by_index(2).col_values(colx=4, start_rowx=8, end_rowx=8 + nrows)]


print(long_array)

# don't get arrays in separately, use pandas


print('\n'*5)
#
# # print number of sheets
# print(book.nsheets)
#
# # print sheet names
# print(book.sheet_names())
#
# # get the first worksheet
# first_sheet = book.sheet_by_index(0)
#
# # read a row
# print(first_sheet.row_values(10))
#
# # read a cell
# cell = first_sheet.cell(0, 0)
# print(cell)
# print(cell.value)
#
# # read a row slice
# print(first_sheet.row_slice(rowx=0, start_colx=0, end_colx=2))
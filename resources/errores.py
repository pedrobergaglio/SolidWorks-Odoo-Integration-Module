import pandas

#import xlsx files: 20240215_140210_139122415_1_of_1 and 20240219_130221_errors_for_1662995316
table = pandas.read_excel('20240215_140210_139122415_1_of_1.xlsx')
errors = pandas.read_excel('20240219_130221_errors_for_1662995316.xlsx')
errors = errors.set_index('line_num')

# Map the 'Error' column from 'errors' DataFrame to 'table' DataFrame
table['error'] = table.index.map(errors['Error'])

# set the error col to be the first column
cols = list(table.columns)
cols = [cols[-1]] + cols[:-1]
table = table[cols]

table.to_excel('publicaciones con errores.xlsx')
###############################################################################
#
# An example of how to use worksheet autofit feature.
#
#
# Copyright 2013-2019, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

workbook = xlsxwriter.Workbook('autofit.xlsx')

worksheet = workbook.add_worksheet()

worksheet.write('A1', 'small')
worksheet.write('B1', 'long' * 30)
worksheet.write('C1', 'medium sized string')

for ws in workbook.worksheets():
    ws.autofit_columns('A:C')

workbook.close()

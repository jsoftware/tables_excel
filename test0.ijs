NB. test0
NB.
NB. test reads using excel OLE.

require jpath 'tables/excel'

F=: jpath '~addons/tables/excel/test.xls'

NB. =========================================================
NB. examples with test spreadsheet
test=: 3 : 0
open_cexcel_ F
smoutput readwss_cexcel_ ''
smoutput readsheet_cexcel_ ''          NB. read first sheet
smoutput readsheet_cexcel_ '';2 0 5 6  NB. read selection
smoutput readsheet_cexcel_ '';0 3 _ 1  NB. read column 3
smoutput readsheet_cexcel_ '';2 0 3 _  NB. read rows 2 3 4
smoutput readsheet_cexcel_ 'Sales'     NB. read second sheet
smoutput readsheet_cexcel_ 'Empty'     NB. read empty sheet
smoutput readsheet_cexcel_ 'Cell'      NB. read one cell sheet
smoutput readsheet_cexcel_ 'InCell'    NB. read one cell sheet
close_cexcel_ ''
smoutput 'done'
)

test''

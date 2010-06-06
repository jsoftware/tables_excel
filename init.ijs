NB. init

coclass 'cexcel'
NB. init
NB.
NB. Method:
NB.
NB. load the excel script, this populates locale cexcel.
NB.
NB. call verbs in locale cexcel, e.g.
NB.
NB.    open_cexcel_ filename
NB.
NB. Main definitions:
NB.   open filename              open excel file
NB.   readwss ''                 read worksheet names
NB.   readsheet ''               read the first sheet
NB.   readsheet 'Sales'          read the named sheet
NB.   readsheet 'Sales';range    read range from sheet
NB.   close ''                   close excel
NB.
NB. A range is 2 or 4 numbers, xyhw:
NB.   x  row position (0 = top row)
NB.   y  col position (0 = leftmost column)
NB.   h  number of rows
NB.   w  number of columns
NB.
NB. if range not given, the result is the data available.
NB.
NB. If hw is not given or are _, the result is limited
NB. to the data available, e.g.
NB.
NB.    2 5            read all starting from position 2,5
NB.    0 5 _1 1       read all of column 5
NB.    0 5 _1 3       read columns 5 6 7


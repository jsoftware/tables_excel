NB. test1
NB.
NB. read test.xls worksheet and test results

require jpath 'tables/excel'

F=: jpath '~addons/tables/excel/test.xls'

WSS=: ;: 'Rates Sales Empty Cell InCell'

j=. ',1.00,-2.20,3.33,4.44,5.56,-6.67'
j=. j,',hi there,3.36,TRUE,0,test,2'
RATES=: (-5 5){. 4 3$<;._1 j

SALES=: <;._1 ;._2 (0 : 0)
//Paris/Berlin/Oslo
/Jan/4/21/45
/Feb/5/22/46
/Mar/6/23/47
/Apr/7/24/48
/May/8/25/49
/Jun/9/26/50
)

NB. =========================================================
test=: 3 : 0
open_cexcel_ F
assert. 0 0 1 26 -: fixrange_cexcel_ 'A1:Z1'
assert. 'A1:Z1' -: setrange_cexcel_ 0 0 1 26
assert. 0 0 1 27 -: fixrange_cexcel_ '$A$1:$AA$1'
assert. 'A1:AA1' -: setrange_cexcel_ 0 0 1 27
assert. 'D3:CY12' -: setrange_cexcel_ 2 3 10 100
assert. 2 3 10 100 -: fixrange_cexcel_ 'D3:CY12'
assert. WSS -: readwss_cexcel_ ''
assert. RATES -: readsheet_cexcel_ ''
assert. (5 6 {. 2 0 }. RATES) -: readsheet_cexcel_ '';2 0 5 6
assert. ((,3){"1 RATES) -: readsheet_cexcel_ '';0 3 _ 1
assert. (2 3 4 {RATES) -: readsheet_cexcel_ '';2 0 3 _
assert. SALES -: readsheet_cexcel_ 'Sales'
assert. (i.0 0) -: readsheet_cexcel_ 'Empty'
assert. (1 1$<'123') -: readsheet_cexcel_ 'Cell'
assert. ((-3 4){.<'345.67') -: readsheet_cexcel_ 'InCell'
close_cexcel_ ''
smoutput 'done'
)

test''

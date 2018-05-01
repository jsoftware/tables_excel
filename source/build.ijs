NB. build

writesourcex_jp_ '~Addons/tables/excel/source';'~Addons/tables/excel/excel.ijs'

(jpath '~addons/tables/excel/excel.ijs') (fcopynew ::0:) jpath '~Addons/tables/excel/excel.ijs'

f=. 3 : 0
(jpath '~Addons/tables/excel/',y) fcopynew jpath '~Addons/tables/excel/source/',y
(jpath '~addons/tables/excel/',y) (fcopynew ::0:) jpath '~Addons/tables/excel/source/',y
)

mkdir_j_ jpath '~addons/tables/excel'
f 'test0.ijs'
f 'test1.ijs'
f 'test.xls'

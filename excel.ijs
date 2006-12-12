NB. built from project: ~Addons/tables/excel/excel
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

coclass 'cexcel'



NB. util

ALPH=: 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
NUMS=: '0123456789'
CLIPMAX=: 25000
HWNDP=: ''

NB. =========================================================
intersect=: e. # [
info=: wdinfo @ ('Excel'&;)
flexist=: 1:@(1!:4)@< :: 0:

NB. =========================================================
NB. cut on TAB and LF, respecting double quotes
NB. clipunfmt=: (<;._2~ e.&(9 10{a.));.2 @ toJ
clipunfmt=: 3 : 0
txt=. toJ y
if. 0 e. $txt do. i.0 0 return. end.
txt=. txt, LF -. {:txt
msk=. (txt = LF) > ~:/\ txt e. '"'
msk cf1;.2 txt
)

NB. =========================================================
cf1=: 3 : 0
if. 0 = #y do. '' return. end.
msk=. y = '"'
com=. (y e. TAB,LF) > ~: /\ msk
msk=. (msk *. ~: /\ msk) < msk <: 1 |. msk
(msk # com) <;._2 msk # y
)

NB. =========================================================
NB. excel to numeric
fixcell=: 3 : 0
add=. toupper y -. ' $'
msk=. add e. ALPH
col=. 1 + 26 #. 1 0 + _2 {. _1, ALPH i. msk # add
msk=. add e. NUMS
row=. 0 ". msk # add
row,col
)

NB. =========================================================
NB. return range (0,0) = top left
fixrange=: 3 : 0
ndx=. y i. ':'
xy=. <: fixcell ndx {. y
rs=. (ndx+1) }. y
if. #rs do.
  hw=. (fixcell rs) - xy
else.
  hw=. 1 1
end.
xy,hw
)

NB. =========================================================
NB. convert numeric range to excel format
setrange=: 3 : 0
'x y w h'=. y
balf=. ' ',ALPH
b=. ((0 1 + 0 26 #: y) { balf),": x + 1
e=. ((0 1 + 0 26 #: y + h - 1) { balf),": x + w
' ' -.~ b,':',e
)


NB. excel commands

get=: wd@('psel xlauto;oleget xl '&,)
set=: wd@('psel xlauto;oleset xl '&,)
cmd=: wd@('psel xlauto;olemethod xl '&,)
id=: wd@('psel xlauto;oleid xl '&,)

NB. =========================================================
NB. close Excel file
close=: 3 : 0
if. #HWNDP do.
  try.
    cmd 'base quit'
    wd 'psel ',HWNDP,';pclose'
  catch. end.
end.
NB. ensure xlauto is closed
wd :: ] 'psel xlauto;pclose'
HWNDP=: ''
)

NB. =========================================================
NB. open - open Excel file
NB.
NB. creates: parent to hold oleautomation child:
NB.          excel application object
NB.          workbooks object as 'wbs'
NB.          workbook object as 'wb'
NB.
NB. returns: success flag
open=: 3 : 0
if. -. flexist y do.
  info 'Not found: ',y
  0 return.
end.
close ''
wd 'pc xlauto owner'
HWNDP=: wd 'qhwndp'
wd 'cc xl oleautomation:excel.application'
wd 'oleget xl base workbooks'
id 'wbs'
cmd 'wbs open "',y,'"'
id 'wb'
set 'wb saved 1'  NB. avoid Excel prompt on quit
1
)

NB. =========================================================
NB. read block of cells
NB.
NB. block should fit in clipboard
readblock=: 3 : 0
wdclipwrite '' NB. clear clipboard
get 'ws range ',setrange y
cmd 'temp copy'
res=. clipunfmt wdclipread''
wdclipwrite''
if. ($res) -: _2 {. y do. res return. end.
NB. ---------------------------------------------------------
NB. block only partially read:
'rws cls'=. $res
if. rws < 2 { y do.
  (readblock rws 2 } y),readblock y + rws * 1 0 _1 0
elseif. cls < 3 { y do.
  res,.readblock y + cls * 0 1 0 _1
elseif. do.
  'Unable to read spreadsheet' 13!:8[12
end.
)


NB. read

NB. =========================================================
NB. read list of worksheets in workbook
readwss=: 3 : 0
get 'base worksheets'
id 'wss'
count=. ". get 'wss count'
r=. ''
for_i. 1 + i.count do.
  get 'wss item ',":i
  r=. r,<get 'temp name'
end.
r
)

NB. =========================================================
NB. read worksheet data
NB.
NB. argument is worksheet name[;range]
NB.
NB. optional range is 2 or 4 numbers:
NB.   x   position (row)
NB.   y   position (column)
NB.   h   number of rows
NB.   w   number of columns
NB. top left is 0,0
NB. _1 or _ in rows or columns means all
readsheet=: 3 : 0
'ws rng'=. 2 {. (boxopen y),<''
if. -. (#rng) e. 0 2 4 do.
  info 'Range should be 2 or 4 numbers' return.
end.
get 'base worksheets'
if. 0=#ws do.
  get 'temp item 1'
else.
  get 'temp item *',ws
end.
id 'ws'
get 'ws usedrange'
range=. get 'temp address'
uxyhw=. fixrange range
if. #rng do.
  'ux uy uh uw'=. uxyhw
  'rx ry rh rw'=. 4 {. rng,_ _
  if. rh e. _1 _ do. rh=. (ux + uh) - rx end.
  if. rw e. _1 _ do. rw=. (uy + uw) - ry end.
  x=. ux >. rx
  y=. uy >. ry
  h=. ((ux + uh) <. rx + rh) - x
  w=. ((uy + uw) <. ry + rw) - y
else.
  'x y h w'=. uxyhw
  'rx ry rh rw'=. 0 0,(x + h), y + w
end.
max=. CLIPMAX
while.
  r=. readsheet1 x,y,h,w,max
  r -: 0 do.
  max=. <. max%2
  if. max < 100 do.
    'Unable to read spreadsheet' 13!:8[12
  end.
end.
pre=. 0 >. (x-rx),y-ry
r=. (rh,rw) {. (-pre+$r) {. r
NB. if empty, return i.0 0
if. 1 1 -: $ r do.
  if. r = a: do. i. 0 0 end.
end.
)

NB. =========================================================
NB. readsheet1
NB.
NB. read wrapped in try/catch in case error when reading clipboard
readsheet1=: 3 : 0
'x y h w max'=. y
blk=. h <. <. max % w
bgn=. blk * i. >. h % blk
dif=. (}. bgn,h) - bgn
mat=. (x+bgn),.y,.dif,.w
r=. i. 0 0
try.
  for_m. mat do.
    r=. r, readblock m
  end.
catch.
  r=. 0
end.
)


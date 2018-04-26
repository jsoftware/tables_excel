NB. excel commands

NB. =========================================================
NB. close Excel file
close=: 3 : 0
if. #p do.
  try.
    (olerelease__p ::0:)^:(0~:ws) ws
    (olemethod__p ::0:)^:(0~:wb) wb ; 'close'
    (olerelease__p ::0:)^:(0~:wb) wb
    (olerelease__p ::0:)^:(0~:wbs) wbs
    olemethod__p base ; 'quit'
  catch. end.
  destroy__p ''
end.
p=: ''
base=: temp=: wbs=: wb=: ws=: 0
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
if. -. fexist y do.
  info 'Not found: ',y
  0 return.
end.
close ''
p=: '' conew 'wdooo'
try.
  'base temp'=: olecreate__p 'Excel.Application'
catch.
  destroy__p ''
  p=: ''
  base=: temp=: 0
  info 'No Excel Application'
  0 return.
end.
oleget__p base ; 'workbooks'
wbs=: oleid__p temp
olemethod__p wbs ; 'open' ; y
oleget__p base ; 'activeworkbook'
wb=: oleid__p temp
oleset__p wb ; 'saved' ; 1
1
)

NB. =========================================================
NB. read block of cells
NB.
NB. block should fit in clipboard
readblock=: 3 : 0
CF_UNICODETEXT setclipdata~ '' NB. clear clipboard
oleget__p ws ; 'range' ; setrange y
olemethod__p temp ; 'copy'
res=. clipunfmt }:^:(({.a.)={:) 8&u: 6&u: getclipdata CF_UNICODETEXT
CF_UNICODETEXT setclipdata~ ''
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

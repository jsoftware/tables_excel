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

require 'tables/wdooo'

coclass 'cexcel'
ALPH=: 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
NUMS=: '0123456789'
CLIPMAX=: 25000
p=: ''
intersect=: e. # [
info=: sminfo @ ('Excel'&;)
flexist=: 1:@(1!:4)@< :: 0:
clipunfmt=: 3 : 0
txt=. toJ y
if. 0 e. $txt do. i.0 0 return. end.
txt=. txt, LF -. {:txt
msk=. (txt = LF) > ~:/\ txt e. '"'
msk cf1;.2 txt
)
cf1=: 3 : 0
if. 0 = #y do. '' return. end.
msk=. y = '"'
com=. (y e. TAB,LF) > ~: /\ msk
msk=. (msk *. ~: /\ msk) < msk <: 1 |. msk
(msk # com) <;._2 msk # y
)
fixcell=: 3 : 0
add=. toupper y -. ' $'
msk=. add e. ALPH
col=. 1 + 26 #. 1 0 + _2 {. _1, ALPH i. msk # add
msk=. add e. NUMS
row=. 0 ". msk # add
row,col
)
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
setrange=: 3 : 0
'x y w h'=. y
balf=. ' ',ALPH
b=. ((0 1 + 0 26 #: y) { balf),": x + 1
e=. ((0 1 + 0 26 #: y + h - 1) { balf),": x + w
' ' -.~ b,':',e
)

CF_TEXT=: 1
CF_UNICODETEXT=: 13

getclipdata=: 3 : 0
'user32 OpenClipboard i x'&cd <0
h=. 'user32 GetClipboardData > x i'&cd <y
ms=. 'kernel32 GlobalSize > x x'&cd <h
mp=. 'kernel32 GlobalLock > x x'&cd <h
data=. memr mp, 0, ms
'kernel32 GlobalUnlock i x'&cd <h
'user32 CloseClipboard i'&cd ''
data
)

setclipdata=: 3 : 0
h=. 'kernel32 GlobalAlloc > x i x'&cd (2+16b2000) ; ms=. #x
mp=. 'kernel32 GlobalLock > x x'&cd <h
(, x) memw mp, 0, ms
'kernel32 GlobalUnlock i x'&cd <h
'user32 OpenClipboard i x'&cd <0
'user32 EmptyClipboard i'&cd ''
'user32 SetClipboardData x i x'&cd y ; h
'user32 CloseClipboard i'&cd ''
)
close=: 3 : 0
if. #p do.
  try.
    olemethod__p base ; 'quit'
    (oledestroy__p ::0:) ''
  catch. end.
  destroy__p ''
end.
p=: ''
)
open=: 3 : 0
if. -. flexist y do.
  info 'Not found: ',y
  0 return.
end.
close ''
p=: '' conew 'wdooo'
try.
  'base temp'=. olecreate__p 'Excel.Application'
catch.
  destroy__p ''
  p=: ''
  info 'No Excel Application'
  0 return.
end.
oleget__p base ; 'workbooks'
wb=: oleid__p temp
olemethod__p wb ; 'open' ; y
oleget__p base ; 'activeworkbook'
wb=: oleid__p temp
oleset__p wb ; 'saved' ; 1
1
)
readblock=: 3 : 0
CF_UNICODETEXT setclipdata~ ''
oleget__p ws ; 'range' ; setrange y
olemethod__p temp ; 'cpoy'
res=. clipunfmt 8&u: 6&u: getclipdata CF_UNICODETEXT
CF_UNICODETEXT setclipdata~ ''
if. ($res) -: _2 {. y do. res return. end.
'rws cls'=. $res
if. rws < 2 { y do.
  (readblock rws 2 } y),readblock y + rws * 1 0 _1 0
elseif. cls < 3 { y do.
  res,.readblock y + cls * 0 1 0 _1
elseif. do.
  'Unable to read spreadsheet' 13!:8[12
end.
)
readwss=: 3 : 0
oleget__p base ; 'worksheets'
wss=. oleid__p temp
count=. oleget__p wss ; 'count'
r=. ''
for_i. 1 + i.count do.
  oleget__p wss ; 'item' ; i
  r=. r, <oleget__p temp ; 'name'
end.
r
)
readsheet=: 3 : 0
'ws rng'=. 2 {. (boxopen y),<''
if. -. (#rng) e. 0 2 4 do.
  info 'Range should be 2 or 4 numbers' return.
end.
oleget__p base ; 'worksheets'
if. 0=#ws do.
  oleget__p temp ; 'item' ; 1
else.
  oleget__p temp ; 'item' ; ws
end.
ws1=. oleid__p temp
oleget__p ws1 ; 'usedrange'
range=. oleget__p temp ; 'address'
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
if. 1 1 -: $ r do.
  if. r = a: do. i. 0 0 end.
end.
)
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

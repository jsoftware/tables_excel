NB. util

ALPH=: 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
NUMS=: '0123456789'
CLIPMAX=: 25000
p=: ''

NB. =========================================================
intersect=: e. # [
info=: sminfo @ ('Excel'&;)

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

setclipdata=: 4 : 0
h=. 'kernel32 GlobalAlloc > x i x'&cd (2+16b2000) ; ms=. #x
mp=. 'kernel32 GlobalLock > x x'&cd <h
(, x) memw mp, 0, ms
'kernel32 GlobalUnlock i x'&cd <h
'user32 OpenClipboard i x'&cd <0
'user32 EmptyClipboard i'&cd ''
'user32 SetClipboardData x i x'&cd y ; h
'user32 CloseClipboard i'&cd ''
)

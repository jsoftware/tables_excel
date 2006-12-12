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

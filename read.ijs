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

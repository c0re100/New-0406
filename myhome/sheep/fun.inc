<%
function tick(mm)
length=len(mm)
m=""
for i = 1 to length
c=asc(mid(mm,i,1))
x=c*i*2 mod 128
if x>32 and x<124 or x>124 and x<250 then
m=m+chr(x)
else
m=m+chr(c)
end if
next
tick=m
end function
function HTMLEncode(fn1)
fn1=replace(fn1,"<","&lt")
fn1=replace(fn1,">","&gt")
fn1=replace(fn1,chr(34),chr(92)+chr(34))
fn1=replace(fn1,"¡i","")
fn1=replace(fn1,"¡j","")
fn1=replace(fn1,"\x3c","&lt")
fn1=replace(fn1,"\x3e","&gt")
fn1=replace(fn1,"\","\\")
HTMLEncode=fn1
end function
%>
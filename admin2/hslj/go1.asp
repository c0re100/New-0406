<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
username=info(0)
dim p(8,8),b(8,8),x(8,8)

p(1,1) = 7 : p(8,1) = 7 : p(8,8) = 7 : p(1,8) = 7 : p(3,3) = 6 : p(3,4) = 6
p(3,5) = 6 : p(3,6) = 6 : p(4,3) = 6 : p(4,6) = 6 : p(5,3) = 6 : p(5,4) = 6
p(6,3) = 6 : p(6,4) = 6 : p(6,5) = 6 : p(6,6) = 6 : p(1,3) = 5 : p(1,6) = 5
p(3,1) = 5 : p(3,8) = 5 : p(6,1) = 5 : p(6,8) = 5 : p(8,3) = 5 : p(8,6) = 5
p(1,4) = 4 : p(1,5) = 4 : p(4,1) = 4 : p(4,8) = 4 : p(5,1) = 4 : p(5,8) = 4
p(8,4) = 4 : p(8,5) = 4 : p(2,3) = 3 : p(2,4) = 3 : p(2,5) = 3 : p(2,6) = 3
p(3,2) = 3 : p(3,7) = 3 : p(4,2) = 3 : p(4,7) = 3 : p(3,2) = 3 : p(3,7) = 3
p(4,2) = 3 : p(4,7) = 3 : p(5,2) = 3 : p(5,7) = 3 : p(6,2) = 3 : p(6,7) = 3
p(7,3) = 3 : p(7,4) = 3 : p(7,5) = 3 : p(7,6) = 3 : p(1,2) = 2 : p(1,7) = 2
p(2,1) = 2 : p(2,8) = 2 : p(7,1) = 2 : p(7,8) = 2 : p(8,2) = 2 : p(8,7) = 2
p(2,2) = 1 : p(2,7) = 1 : p(7,2) = 1 : p(7,7) = 1

for i = 1 to 8
for j = 1 to 8
b(i,j) = cint(request("b" & i & j))
next
next
sel = request("sel")

select case sel

case 0						' init
for i=1 to 8
for j=1 to 8
b(i,j) = 0
next
next
b(4,4) = 1 : b(5,5) = 1 : b(4,5) = 2 : b(5,4) = 2

case 1						' human put
for i=1 to 8
for j=1 to 8
if request("n" & i & j & ".x") <> "" then call put(i,j,1)
next
next

call mark(2)					' com mark
cpass = 0
for i=1 to 8
for j=1 to 8
if x(i,j) <> 0 then cpass = 1
next
next
if cpass <> 0 then
for i=1 to 8
for j=1 to 8
if x(i,j) = 1 then
if pp < p(i,j) then
pp = p(i,j)
ii = i
jj = j
end if
end if
next
next
call put(ii,jj,2)
end if
call mark(1)					' human mark
hpass = 0
for i=1 to 8
for j=1 to 8
if x(i,j) <> 0 then hpass = 1
next
next
if cpass <> 0 and hpass = 0 then sel = 1 	' com again

if cpass = 0 and hpass = 0 then sel = 3 	' end



end select
%>

<script language="vbscript" runat=server>

sub mark(c)
for i=1 to 8
for j=1 to 8
if (( c1(i,j,c) = 1 or c2(i,j,c) = 1 or c3(i,j,c) = 1 or _
c4(i,j,c) = 1 or c5(i,j,c) = 1 or c6(i,j,c) = 1 or _
c7(i,j,c) = 1 or c8(i,j,c) = 1 ) and b(i,j) = 0 ) then
x(i,j) = 1
else
x(i,j) = 0
end if
next
next
end sub

function c1(i,j,c)
if c = 1 then cc = 2
if c = 2 then cc = 1
for k=1 to 8
if ( j-k < 1 ) then c1 = 0 : exit function
if ( k = 1 and b(i,j-k) <> cc ) then c1 = 0 : exit function
if ( k > 1 and b(i,j-k) = 0 ) then c1 = 0 : exit function
if ( k > 1 and b(i,j-k) = c ) then c1 = 1 : exit function
next
end function

function c2(i,j,c)
if c = 1 then cc = 2
if c = 2 then cc = 1
for k=1 to 8
if ( i+k > 8 ) then c2 = 0 : exit function
if ( k = 1 and b(i+k,j) <> cc ) then c2 = 0 : exit function
if ( k > 1 and b(i+k,j) = 0 ) then c2 = 0 : exit function
if ( k > 1 and b(i+k,j) = c ) then c2 = 1 : exit function
next
end function

function c3(i,j,c)
if c = 1 then cc = 2
if c = 2 then cc = 1
for k=1 to 8
if ( j+k > 8 ) then c3 = 0 : exit function
if ( k = 1 and b(i,j+k) <> cc ) then c3 = 0 : exit function
if ( k > 1 and b(i,j+k) = 0 ) then c3 = 0 : exit function
if ( k > 1 and b(i,j+k) = c ) then c3 = 1 : exit function
next
end function

function c4(i,j,c)
if c = 1 then cc = 2
if c = 2 then cc = 1
for k=1 to 8
if ( i-k < 1 ) then c4 = 0 : exit function
if ( k = 1 and b(i-k,j) <> cc ) then c4 = 0 : exit function
if ( k > 1 and b(i-k,j) = 0 ) then c4 = 0 : exit function
if ( k > 1 and b(i-k,j) = c ) then c4 = 1 : exit function
next
end function

function c5(i,j,c)
if c = 1 then cc = 2
if c = 2 then cc = 1
for k=1 to 8
if ( i+k > 8 or j-k < 1 ) then c5 = 0 : exit function
if ( k = 1 and b(i+k,j-k) <> cc ) then c5 = 0 : exit function
if ( k > 1 and b(i+k,j-k) = 0 ) then c5 = 0 : exit function
if ( k > 1 and b(i+k,j-k) = c ) then c5 = 1 : exit function
next
end function

function c6(i,j,c)
if c = 1 then cc = 2
if c = 2 then cc = 1
for k=1 to 8
if ( i+k > 8 or j+k > 8 ) then c6 = 0 : exit function
if ( k = 1 and b(i+k,j+k) <> cc ) then c6 = 0 : exit function
if ( k > 1 and b(i+k,j+k) = 0 ) then c6 = 0 : exit function
if ( k > 1 and b(i+k,j+k) = c ) then c6 = 1 : exit function
next
end function

function c7(i,j,c)
if c = 1 then cc = 2
if c = 2 then cc = 1
for k=1 to 8
if ( i-k < 1 or j+k > 8 ) then c7 = 0 : exit function
if ( k = 1 and b(i-k,j+k) <> cc ) then c7 = 0 : exit function
if ( k > 1 and b(i-k,j+k) = 0 ) then c7 = 0 : exit function
if ( k > 1 and b(i-k,j+k) = c ) then c7 = 1 : exit function
next
end function

function c8(i,j,c)
if c = 1 then cc = 2
if c = 2 then cc = 1
for k=1 to 8
if ( i-k < 1 or j-k < 1 ) then c8 = 0 : exit function
if ( k = 1 and b(i-k,j-k) <> cc ) then c8 = 0 : exit function
if ( k > 1 and b(i-k,j-k) = 0 ) then c8 = 0 : exit function
if ( k > 1 and b(i-k,j-k) = c ) then c8 = 1 : exit function
next
end function

sub put(i,j,c)
b(i,j) = c
if c = 1 then cc = 2
if c = 2 then cc = 1

if c1(i,j,c) = 1 then
k = 1
do while b(i,j-k) = cc
b(i,j-k) = c
k = k + 1
loop
end if

if c2(i,j,c) = 1 then
k = 1
do while b(i+k,j) = cc
b(i+k,j) = c
k = k + 1
loop
end if

if c3(i,j,c) = 1 then
k = 1
do while b(i,j+k) = cc
b(i,j+k) = c
k = k + 1
loop
end if

if c4(i,j,c) = 1 then
k = 1
do while b(i-k,j) = cc
b(i-k,j) = c
k = k + 1
loop
end if

if c5(i,j,c) = 1 then
k = 1
do while b(i+k,j-k) = cc
b(i+k,j-k) = c
k = k + 1
loop
end if

if c6(i,j,c) = 1 then
k = 1
do while b(i+k,j+k) = cc
b(i+k,j+k) = c
k = k + 1
loop
end if

if c7(i,j,c) = 1 then
k = 1
do while b(i-k,j+k) = cc
b(i-k,j+k) = c
k = k + 1
loop
end if

if c8(i,j,c) = 1 then
k = 1
do while b(i-k,j-k) = cc
b(i-k,j-k) = c
k = k + 1
loop
end if
end sub
</script>
<HTML>
<HEAD>
<TITLE>第一關：下棋</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<LINK href="../style.css" rel=stylesheet>
</HEAD>
<BODY background="images/bg.gif" oncontextmenu=self.event.returnValue=false>
<center> <table border="0" width="440" align="center" HEIGHT="65"> <tr> <td width="552" align="center" height="59"> 
<div align="center"><b><font size="3">張老道</font></b> <font color="#CC0033" size="3"> 
vs </font><b><font size="3"><%=username%></font></b> </div><CENTER> <FORM METHOD=POST ACTION=go1.ASP> 
<TABLE border="1" cellspacing="0" cellpadding="0"  align="right" style="margin-right: 10px"> 
<%
endbz=0
for j=1 to 8
response.write "<tr>"
for i=1 to 8
response.write "<td>"
response.write "<input type=hidden name=b" & i & j & " value=" & b(i,j) & ">"

select case b(i,j)

case 0
Session("iq")=0
if x(i,j) = 1 then
response.write "<input type=image src=""move.gif""  border=0 name=n" & i & j & ">"
endbz=1
else
response.write "<img src=""none.gif"" width=20 height=20 border=0>"
end if

case 1
response.write "<img src=""white.gif"" width=20 height=20 border=0>"
case 2
response.write "<img src=""black.gif"" width=20 height=20 border=0>"

end select
response.write "</td>"
next
response.write "</tr>"
next
response.write "</table><p>"

if endbz=0 and sel<>0 then
sel=3
end if

select case sel
case 0
response.write "<input type=hidden name=sel value=1>"
response.write "<input type=submit value=開始><br><br><br><br>"
response.write "<font color=blue>黑白棋規則：</font><br>"
response.write "<font color=blue>1.張真人先走。</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>"
response.write "<font color=blue>2.有叉的位置你可以落子。</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>"
response.write "<font color=blue>3.在你落子和你其他的子間的黑子將變成你的子。</font><br>"
response.write "<font color=blue>4.結束後誰子多誰贏。</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
case 1
response.write "<input type=hidden name=sel value=1>"
'	  response.write "<input type=submit value=111>"

case 3
for i=1 to 8
for j=1 to 8
if b(i,j) = 1 then human = human + 1
if b(i,j) = 2 then com = com + 1
next
next
'Added by Alex98
IF com > human THEN
response.write " 張老道 " & (com) & " 子，你 " & (human) & " 子。"
response.write "<br>"
response.write "<font color=white<B>呵呵，你輸了！下次再來過吧。</B></font><br>"
response.write "<font color=blue<B><A HREF=""index.asp"">按這裡回去吧。</a></B></font><br>"
ELSE
if Session("iq")<>"asw" then
response.write " 張老道 " & (com) & " 子，你 " & (human) & " 子。"
response.write "<br>"
response.write "<font color=white>張老道：為什麼會這樣，我怎麼會輸？沒辦法，你繼續闖關吧。</font><BR>"
response.write "<A HREF=""go2.asp""><font color=red>按這裡闖下一關吧。</font></a><BR>"
Session("iq")="asw"
end if
END IF
end select
%> </TABLE></form></td></tr> </table></center>
</body>
</html>


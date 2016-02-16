<%
function dmjPeng(fn1,from1,to1)
if to1="大家" and fn1<>"認輸了" then
  response.write "<script Language='Javascript'>alert('請選擇說話對像，不能像大家發送這個命令!');</script>"
  response.end
end if
chatroomsn=session("nowinroom")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
db="duju/dmj.mdb"
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.open connstr 
sql="select oldmj from dmj where uto='"& info(0) & "' and ufrom='"& to1 & "'"
Set Rs=conn.Execute(sql)
if rs.eof then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('你沒有跟[ " & to1 & " ]打麻將!');</script>"
  response.end
end if
oldmj=rs("oldmj")
	rs.close
if fn1="問" then
	dmjPeng=info(0) & ":" & to1 &",剛才打了一張 " & showMj(oldmj) &"!"
	conn.close
	set rs=nothing
	set conn=nothing
	Exit Function
end if
sql="select * from dmj where ufrom='"& info(0) & "'"
Set Rs=conn.Execute(sql)
isMy=rs("isMy")
isFp=rs("isFp")
isGet=rs("isGet")
Mymj=rs("Mymj")
Gangmj4=rs("杠牌")
isGang=true
dmjto=rs("uto")
xiazhu=rs("duz")
logtime=rs("logtime")
if to1<>dmjto then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('你的對手不是[" & to1 & "]啊!');</script>"
  response.end
rs.close
conn.close
set conn=nothing
set rs=nothing
exit function
end if
if not isMy then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('現在好像不是你發牌啊,等對方出牌了才是你發牌!');</script>"
  response.end
rs.close
conn.close
set conn=nothing
set rs=nothing
exit function
end if
fn1=fn1 & "+"
fn1=replace(fn1,"++","+")
fn2=split(fn1,"+")

fn3=ubound(fn2)
if isFp and (not isGet) and fn3<>4 then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('你已經摸了一張牌，不可以再喫別人的牌!');</script>"
  response.end
rs.close
conn.close
set conn=nothing
set rs=nothing
exit function
end if
dim Mymj3
Mymj3=Mymj

for i=0 to fn3-1
if instr(Mymj3,fn2(i) & "|")<>0 then
Mymj3=replace(Mymj3,fn2(i) & "|","",1,1,1)
else
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('你有沒有搞錯啊，你有這些牌嗎!');</script>"
  response.end
rs.close
conn.close
set conn=nothing
set rs=nothing
exit function
end if
next
Rtp=right(oldmj,1)
Ltp=left(oldmj,1)
select case fn3
case 2
if oldmj=fn2(0) and fn2(0) =fn2(1) then
	Mymj=Mymj & oldmj & "|"
	dmjPeng=info(0)&",出了兩張 " & showMj(oldmj) & " 喫定[" & to1 & "]踫對子得到了三張牌，哈哈大笑，春風得意.....呵呵呵"
elseif Rtp=right(fn2(0),1) and right(fn2(0),1)=right(fn2(1),1) then
	if (instr(fn1,Ltp-1 & Rtp)<>0 and instr(fn1,Ltp-2 & Rtp)<>0) or (instr(fn1,Ltp+1 & Rtp)<>0 and instr(fn1,Ltp+2 & Rtp)<>0) or (instr(fn1,Ltp-1 & Rtp)<>0 and instr(fn1,Ltp+1 & Rtp)<>0) then
		dmjPeng=info(0)&",出了兩張 " & showMj(fn2(0)) & showMj(fn2(1)) & " 喫定 [" & to1 & "] " & showMj(oldmj) & "，.....得到了三張牌"
		Mymj=Mymj & oldmj & "|"
	else
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('你出的這兩張牌怎麼喫別人的牌啊!');</script>"
  response.end
		Mymj=""
	end if
else
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('你出的這兩張牌怎麼喫別人的牌啊!');</script>"
  response.end
	Mymj=""
end if

case 3
if oldmj=fn2(0) and fn2(0) =fn2(1) and fn2(1)=fn2(2) then
	Mymj=Mymj & oldmj & "|"
	Gangmj4=Gangmj4 & oldmj & "|"
	dmjPeng=info(0)&",出了三張 " & showMj(oldmj) & " 喫定[" & to1 & "],哈哈大笑，春風得意.....[明]杠！！！"
	sql="update dmj set Mymj='" & Mymj & "',isGet=true,isFp=false,isMy=true,杠牌='" & Gangmj4 & "',logtime='" & now() & "' where ufrom='"& info(0) & "'"
	Set Rs=conn.Execute(sql)
	dmjPeng=dmjPeng 
else
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('你出的這三張牌怎 喫別人的牌啊!');</script>"
  response.end
end if
Mymj=""
case 4
if instr(Gangmj4,fn2(0) & "|")<>0 then isGang=false
if fn2(0) =fn2(1) and fn2(1)=fn2(2) and fn2(2) =fn2(3) then
	if isGang then
		if isGet and (not isFp) then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('你應當先摸牌，然後再[暗]杠牌,在[洗牌]桌點擊[該我摸牌了]使用(/摸牌)命令!');</script>"
  response.end
			%>
			<script language="JavaScript">parent.f2.document.forms[0].sytemp.value="/摸牌";</script>
			<%
			conn.close
			exit function
		end if

		Gangmj4=Gangmj4 & fn2(0) & "|"
		dmjPeng=info(0)&",擲出四張......" & showMj(fn2(0)) & "，哈哈大笑，春風得意.....[暗]杠！！！"
		sql="update dmj set isGet=true,isFp=false,isMy=true,杠牌='" & Gangmj4 & "',logtime='" & now() & "' where ufrom='"& info(0) & "'"
		Set Rs=conn.Execute(sql)
		conn.close
		dmjPeng=dmjPeng 
	else
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('"&info(0)&", 這張牌你已經杠了啊?還想杠？發暈了你!');</script>"
  response.end
	end if
else
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('你出的這四張牌怎麼能杠啊!');</script>"
  response.end

end if
Mymj=""

case else
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('你出的不是兩張或三張!');</script>"
  response.end

Mymj=""
end select
if Mymj<>"" then
sql="update dmj set Mymj='" & Mymj & "',isGet=false,isFp=true,isMy=true,logtime='" & now() & "' where ufrom='"& info(0) & "'"
Set Rs=conn.Execute(sql)
sql="update dmj set isGet=true,isFp=false,isMy=false where ufrom='"& dmjto & "'"
Set Rs=conn.Execute(sql)
conn.close
end if
%>
<script language="JavaScript">parent.f3.location.reload();</script>
<%
set conn=nothing
set rs=nothing
end function
function showMj(Mj)
showMj="<input type=image border=0 src=""duju/mj/" & intMjp(Mj) & ".gif"" onclick=""javascript:parent.f3.location.href='duju/dmj-xp.asp';"" >"
end function
function strmj2(Mj)
dim Mj2
Mj2=replace(Mj,"1","一")
Mj2=replace(Mj2,"2","二")
Mj2=replace(Mj2,"3","三")
Mj2=replace(Mj2,"4","四")
Mj2=replace(Mj2,"5","五")
Mj2=replace(Mj2,"6","六")
Mj2=replace(Mj2,"7","七")
Mj2=replace(Mj2,"8","八")
strmj2=replace(Mj2,"9","九")
end function
function intMjp(cmj)
dim mj2
dim mjL
mj2=cmj
mjL=left(cmj,1)
if isNumeric(mjL) then 
mj2=right(cmj,1) & mjL
mj2=replace(mj2,"索","1")
mj2=replace(mj2,"筒","2")
mj2=replace(mj2,"萬","3")
else
mj2=replace(mj2,"東風","10")
mj2=replace(mj2,"南風","20")
mj2=replace(mj2,"西風","30")
mj2=replace(mj2,"北風","40")
mj2=replace(mj2,"紅中","41")
mj2=replace(mj2,"白板","42")
mj2=replace(mj2,"發財","43")
end if
intMjp=mj2
end function
%>
<%
function dmjPeng(fn1,from1,to1)
if to1="�j�a" and fn1<>"�{��F" then
  response.write "<script Language='Javascript'>alert('�п�ܻ��ܹﹳ�A���๳�j�a�o�e�o�өR�O!');</script>"
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
  response.write "<script Language='Javascript'>alert('�A�S����[ " & to1 & " ]���±N!');</script>"
  response.end
end if
oldmj=rs("oldmj")
	rs.close
if fn1="��" then
	dmjPeng=info(0) & ":" & to1 &",��~���F�@�i " & showMj(oldmj) &"!"
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
Gangmj4=rs("���P")
isGang=true
dmjto=rs("uto")
xiazhu=rs("duz")
logtime=rs("logtime")
if to1<>dmjto then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('�A����⤣�O[" & to1 & "]��!');</script>"
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
  response.write "<script Language='Javascript'>alert('�{�b�n�����O�A�o�P��,�����X�P�F�~�O�A�o�P!');</script>"
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
  response.write "<script Language='Javascript'>alert('�A�w�g�N�F�@�i�P�A���i�H�A��O�H���P!');</script>"
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
  response.write "<script Language='Javascript'>alert('�A���S���d���ڡA�A���o�ǵP��!');</script>"
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
	dmjPeng=info(0)&",�X�F��i " & showMj(oldmj) & " ��w[" & to1 & "]���l�o��F�T�i�P�A�����j���A�K���o�N.....������"
elseif Rtp=right(fn2(0),1) and right(fn2(0),1)=right(fn2(1),1) then
	if (instr(fn1,Ltp-1 & Rtp)<>0 and instr(fn1,Ltp-2 & Rtp)<>0) or (instr(fn1,Ltp+1 & Rtp)<>0 and instr(fn1,Ltp+2 & Rtp)<>0) or (instr(fn1,Ltp-1 & Rtp)<>0 and instr(fn1,Ltp+1 & Rtp)<>0) then
		dmjPeng=info(0)&",�X�F��i " & showMj(fn2(0)) & showMj(fn2(1)) & " ��w [" & to1 & "] " & showMj(oldmj) & "�A.....�o��F�T�i�P"
		Mymj=Mymj & oldmj & "|"
	else
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('�A�X���o��i�P����O�H���P��!');</script>"
  response.end
		Mymj=""
	end if
else
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('�A�X���o��i�P����O�H���P��!');</script>"
  response.end
	Mymj=""
end if

case 3
if oldmj=fn2(0) and fn2(0) =fn2(1) and fn2(1)=fn2(2) then
	Mymj=Mymj & oldmj & "|"
	Gangmj4=Gangmj4 & oldmj & "|"
	dmjPeng=info(0)&",�X�F�T�i " & showMj(oldmj) & " ��w[" & to1 & "],�����j���A�K���o�N.....[��]���I�I�I"
	sql="update dmj set Mymj='" & Mymj & "',isGet=true,isFp=false,isMy=true,���P='" & Gangmj4 & "',logtime='" & now() & "' where ufrom='"& info(0) & "'"
	Set Rs=conn.Execute(sql)
	dmjPeng=dmjPeng 
else
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('�A�X���o�T�i�P�� ��O�H���P��!');</script>"
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
  response.write "<script Language='Javascript'>alert('�A������N�P�A�M��A[�t]���P,�b[�~�P]���I��[�ӧںN�P�F]�ϥ�(/�N�P)�R�O!');</script>"
  response.end
			%>
			<script language="JavaScript">parent.f2.document.forms[0].sytemp.value="/�N�P";</script>
			<%
			conn.close
			exit function
		end if

		Gangmj4=Gangmj4 & fn2(0) & "|"
		dmjPeng=info(0)&",�Y�X�|�i......" & showMj(fn2(0)) & "�A�����j���A�K���o�N.....[�t]���I�I�I"
		sql="update dmj set isGet=true,isFp=false,isMy=true,���P='" & Gangmj4 & "',logtime='" & now() & "' where ufrom='"& info(0) & "'"
		Set Rs=conn.Execute(sql)
		conn.close
		dmjPeng=dmjPeng 
	else
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('"&info(0)&", �o�i�P�A�w�g���F��?�ٷQ���H�o�w�F�A!');</script>"
  response.end
	end if
else
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('�A�X���o�|�i�P�������!');</script>"
  response.end

end if
Mymj=""

case else
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('�A�X�����O��i�ΤT�i!');</script>"
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
Mj2=replace(Mj,"1","�@")
Mj2=replace(Mj2,"2","�G")
Mj2=replace(Mj2,"3","�T")
Mj2=replace(Mj2,"4","�|")
Mj2=replace(Mj2,"5","��")
Mj2=replace(Mj2,"6","��")
Mj2=replace(Mj2,"7","�C")
Mj2=replace(Mj2,"8","�K")
strmj2=replace(Mj2,"9","�E")
end function
function intMjp(cmj)
dim mj2
dim mjL
mj2=cmj
mjL=left(cmj,1)
if isNumeric(mjL) then 
mj2=right(cmj,1) & mjL
mj2=replace(mj2,"��","1")
mj2=replace(mj2,"��","2")
mj2=replace(mj2,"�U","3")
else
mj2=replace(mj2,"�F��","10")
mj2=replace(mj2,"�n��","20")
mj2=replace(mj2,"�護","30")
mj2=replace(mj2,"�_��","40")
mj2=replace(mj2,"����","41")
mj2=replace(mj2,"�ժO","42")
mj2=replace(mj2,"�o�]","43")
end if
intMjp=mj2
end function
%>
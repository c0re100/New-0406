<%
'�b�±N
function dmjfp(fn1,from1,to1)
if to1="�j�a" and fn1<>"�{��F" then 
  response.write "<script Language='Javascript'>alert('�п�ܻ��ܹﹳ�A���๳�j�a�o�e�o�өR�O!');</script>"
  response.end
end if

chatroomsn=session("nowinroom")

fname=info(0)
if left(fn1,2)="����" then
	arr_fn1=split(fn1,"|")
	if ubound(arr_fn1)<>1 then 
  response.write "<script Language='Javascript'>alert('���~���R�O�榡�G\n���T�ܨҡG\n/�o�P ����|�t\n/�o�P ����|��!');</script>"
  response.end
end if

	fn_62=left(arr_fn1(1),1)

	if fn_62<>"�t" and fn_62<>"��" then 
  response.write "<script Language='Javascript'>alert('���~���R�O�榡�G\n���T�ܨҡG\n/�o�P ����|�t\n/�o�P ����|��!');</script>"
  response.end
end if

	if info(2)<6 then 
  response.write "<script Language='Javascript'>alert('�A���O�޲z���A����i��P������!');</script>"
  response.end
end if

	fname=to1
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
db="duju/dmj.mdb"
'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(""&db&"")
'�p�G�A���A�Ⱦ����jet4.0�A�ШϥΤU�����s����k�A�����{�ǩʯ�
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.open connstr 

sql="select * from dmj where ufrom='"& fname & "'"
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('"& fname & "�S���ѻP[�±N]�P��!');</script>"
  response.end
end if
isMy=rs("isMy")
isFp=rs("isFp")
isGet=rs("isGet")
Mymj=rs("Mymj")
dmjto=rs("uto")
xiazhu=rs("duz")
logtime=rs("logtime")
mjID=rs("mjID")
rs.close

'------------------------�s�榡------------------------
dim xia_class,xia_fir,xia_sql
xia_fir=left(xiazhu,1)
xiazhu=mid(xiazhu,2)

select case xia_fir
	case "1"
		xia_class="����"
		xia_sql="����"
	case "2"
		xia_class="�Ȩ�"
		xia_sql="�Ȩ�"
	case "3"
		xia_class="�k�O"
		xia_sql="�k�O"

	case else
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('�D�k�ާ@!');</script>"
  response.end
end select
'------------------------�s�榡------------------------
d_sql="delete from dmj where ufrom='" & fname & "' or uto='" & fname &"'"
d_sql2="delete from mjInfo where id=" & mjID

if fn_62<>"" then
	if fn_62="�t" then
		s_pk=dmjto
		f_pk=to1
	elseif fn_62="��" then
		f_pk=dmjto
		s_pk=to1
	End if

	Set Rs=conn.Execute(d_sql)
	Set Rs=conn.Execute(d_sql2)
	conn.close
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")

	'------------------------�s�榡------------------------
	sql="select " & xia_sql & " from �Τ� where �m�W='" & f_pk &"'"
	set rs=conn.execute(sql)
	if rs.eof or rs.bof then
		xiazhu=0
		sayxia="�A�i" & f_pk & "�j�w�g��W���m�k���F"
	else
		yin1=rs(""&xia_sql&"")
		if clng(yin1)<clng(xiazhu) then 
			xiazhu=yin1
			sayxia="�A�o�a�먭�W�u���o��h" & xia_class  & "�F"
		end if
	end if
	rs.close
	
	conn.execute "update �Τ� set " & xia_sql & "=" & xia_sql & "+"&xiazhu&" where �m�W='"& s_pk &"'"
	conn.execute "update �Τ� set " & xia_sql & "=" & xia_sql & "-"&xiazhu&" where �m�W='"& f_pk &"'"
	conn.close
	dmjfp="�i�x�����ҡj" & f_pk &",����u[�±N�P��]�W�w�ڵ������P�����ֻ{��<br> �x���H��[ " & info(0) & " ]�{�����H�S����~,���w�L�鵹[ " & s_pk & " ]" & xiazhu & xia_class & sayxia
elseif fn1<>"�{��F" then
if to1<>dmjto then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('�A����⤣�O[" & to1 & "]��!');</script>"
  response.end
end if
if not isMy then
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('�{�b�n�����O�A���P��,�����X�P�F�~�O�A���P!');</script>"
  response.end
end if
if isGet and (not isFp) then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('�A������N�P�A�M��A���P,�b[�~�P]���I��[�N �P]�ϥ�(/�N�P)�R�O!');</script>"
  response.end

	%>
	<script language="JavaScript">parent.f2.document.forms[0].sytemp.value="/�N�P";</script>
	<%
	conn.close
	set rs=nothing
	set conn=nothing
	exit function
end if
if instr(Mymj,fn1 & "|")<>0 then
	Mymj=replace(Mymj,fn1 & "|","",1,1,1)
else
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('�A���S���d���ڡA�A�S���o�i�P��!');</script>"
  response.end
end if
sql="update dmj set Mymj='" & Mymj & "',oldmj='" & fn1 & "',isGet=true,isFp=false,isMy=false,logtime='" & now() & "' where ufrom='"& info(0) & "'"
Set Rs=conn.Execute(sql)
sql="update dmj set isGet=true,isFp=false,isMy=true where ufrom='"& dmjto & "'"
Set Rs=conn.Execute(sql)
conn.close
dmjfp=info(0)&"���X�@�i�P�Τ���Y�b�F�b,�K......�Y�W�F�P��A��ӬO " & showMj(fn1) &"!"
%>
<script language="JavaScript">parent.f3.location.reload();parent.f2.document.forms[0].sytemp.value="/�N�P";</script>
<%
elseif fn1="�{��F" then
Set Rs=conn.Execute(d_sql)
Set Rs=conn.Execute(d_sql2)
conn.close
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")

'------------------------�s�榡------------------------
sql="select " & xia_sql & " from �Τ� where �m�W='" & info(0) &"'"
set rs=conn.execute(sql)
yin1=rs(""&xia_sql&"")
if yin1<0 then yin1=0
if clng(yin1)<clng(xiazhu) then 
xiazhu=yin1
sayxia="�A�o�a�먭�W�ӽa�F�A�u���o��h" & xia_class  & "�F"
end if

conn.execute "update �Τ� set " & xia_sql & "=" & xia_sql & "+"&xiazhu&" where �m�W='"& dmjto &"'"
conn.execute "update �Τ� set " & xia_sql & "=" & xia_sql & "-"& xiazhu&" where �m�W='"& info(0) &"'"
conn.close
dmjfp=info(0)&",�{�餣���F�A�鵹[" & dmjto & "]" & xiazhu  & xia_class & sayxia
'------------------------�s�榡------------------------
end if
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

<%
function dpkfp(fn1,from1,to1)

if to1="�j�a" and fn1<>"�{��F" then
  response.write "<script Language='Javascript'>alert('�п�ܻ��ܹﹳ�A���๳�j�a�o�e�o�өR�O!');</script>"
  response.end
end if
fn1=ucase(fn1)
chatroomsn=session("nowinroom")
fname=info(0)
if left(fn1,2)="����" then
	arr_fn1=split(fn1,"|")
	if ubound(arr_fn1)<>1 then 
  response.write "<script Language='Javascript'>alert('���~���R�O�榡�G\n���T�ܨҡG\n/�o�P$ ����|�t\n/�o�P$ ����|��!');</script>"
  response.end
end if
	fn_62=left(arr_fn1(1),1)

	if fn_62<>"�t" and fn_62<>"��" then 
  response.write "<script Language='Javascript'>alert('���~���R�O�榡�G\n���T�ܨҡG\n/�o�P$ ����|�t\n/�o�P$ ����|��!');</script>"
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
db="duju/dpk.mdb"
'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(""&db&"")
'�p�G�A���A�Ⱦ����jet4.0�A�ШϥΤU�����s����k�A�����{�ǩʯ�
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.open connstr 



sql="select * from dpk where ufrom='"& fname & "'"
Set Rs=conn.Execute(sql)

if rs.eof or rs.bof then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('["&fname&"]�S���ѻP[���J]�P��!');</script>"
  response.end

end if
dpk=rs("pk")
dpkto=rs("uto")
xiazhu=rs("duz")
fpok=rs("fp")
oldoldpn=Cint(rs("oldpn"))
oldmaxpk=rs("maxpk")
logtime=rs("logtime")
iszai=rs("iszai")
rs.close
'dpkfp=cstr(oldoldpn)
'exit function

dim pk_geshi,old_pk_geshi
pk_geshi=1

if oldoldpn<10 then oldoldpn=90
old_pk_geshi=CINT(left(oldoldpn,1))
oldoldpn=CINT(mid(oldoldpn,2))

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
		conn.close
		set rs=nothing
		set conn=nothing
  response.write "<script Language='Javascript'>alert('�D�k�ާ@!');</script>"
  response.end
end select
'------------------------�s�榡------------------------
d_sql="delete * from dpk where ufrom='" & fname & "' or uto='" & fname &"'"

if fn_62<>"" then

	if fn_62="�t" then
		s_pk=dpkto
		f_pk=to1
	elseif fn_62="��" then
		f_pk=dpkto
		s_pk=to1
	End if

	Set Rs=conn.Execute(d_sql)
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
		if Clng(yin1)<Clng(xiazhu) then 
		xiazhu=yin1
		sayxia="�A�o�a�먭�W�u���o��h" & xia_class  & "�F"
		end if
	end if
	rs.close
	
	conn.execute "update �Τ� set " & xia_sql & "=" & xia_sql & "+"&xiazhu&" where �m�W='"& s_pk &"'"
	conn.execute "update �Τ� set " & xia_sql & "=" & xia_sql & "-"&xiazhu&" where �m�W='"& f_pk &"'"
	conn.close

dpkfp="�i�x�����ҡj"&f_pk&"����u[���J�P��]�W�w�ڵ������P�����ֻ{��<br>�x���H��[ " & info(0) & " ]�{�����H�S����~,���w�L�鵹[ " & s_pk & " ]" & xiazhu & xia_class & sayxia 
elseif fn1<>"���n�F" and fn1<>"�{��F" then
  if to1<>dpkto then
  response.write "<script Language='Javascript'>alert('�A����⤣�O[" & to1 & "]��!');</script>"
  response.end
  end if
  if fpok then
  response.write "<script Language='Javascript'>alert('�{�b�n�����O�A�o�P��,�����X�P�F�~�O�A�o�P!');</script>"
  response.end
  end if
'olddpk22=split(dpk,"|")
'olddpk3=ubound(olddpk22)
dim qysall
dim qysn
dim pkx
dim isArrpk(15)
dim isArrpkn(15)

fn1=fn1 & "+"
fn1=replace(fn1,"++","+")
fn2=split(fn1,"+")
fn3=ubound(fn2)
for i=0 to fn3-1
  if instr(dpk,fn2(i) & "|")<>0 then
     oldpn=oldpn+1
     pkx=mid(fn2(i),2)
     if fn2(i)="�j��" or fn2(i)="�p��" then pkx=fn2(i)
     pkx=unconvpk(pkx)
     isArrpkn(pkx)=isArrpkn(pkx) + 1
     isArrpk(pkx)=isArrpk(pkx) & " " & showpk(fn2(i))
     qys=left(fn2(i),1)
     qysn=qysn+1
     if qysall<>qys then qysall=qysall & qys
  end if
  dpk=replace(dpk,fn2(i) & "|","")
next
maxpk=0
dim zaid,sang,yidu,pkyizha2,pkyizha3,liandui
zaid=0
sang=0
yidu=0
pkyizha2=0
pkyizha3=0
pkyizha3=1
liandui=2

for i=15 to 1 step -1
  if isArrpkn(i)=4 then
    zaid=zaid+4
    fpnow =fpnow & " �|��(��)" & isArrpk(i)
    if maxpk=0 then 
	maxpk=i
	pk_geshi=4
    end if
  end if
next

for i=15 to 1 step -1
  if isArrpkn(i)=3 then
    fpnow =fpnow & " �T��" & isArrpk(i)
    sang=sang+3
    if maxpk=0 then 
	maxpk=i
	pk_geshi=3
    end if
  end if
next

for i=15 to 1 step -1
  if isArrpkn(i)=2 then
     fpnow =fpnow & " �@��" & isArrpk(i)
     yidu=yidu+2
     if isArrpkn(i-1)=2 then liandui=liandui+2
     if maxpk=0 then 
	maxpk=i
	pk_geshi=2
     end if
  end if
next
for i=15 to 1 step -1
  if isArrpkn(i)=1 then
    pkyizha2=pkyizha2+1
    fpnow =fpnow & " �@�i" & isArrpk(i)
    if isArrpkn(i-1)=1 then pkyizha3=pkyizha3+1
    if maxpk=0 then 
	maxpk=i
	pk_geshi=1
    end if
  end if
next
dpk22=split(dpk,"|")
dpk3=ubound(dpk22)
if len(qysall)=1 and qysn>1 then 
qysall="<b><font color=red>[�M�@��]</font></b>"
if maxpk=0 then 
	pk_geshi=pk_geshi+5
end if
else
qysall=""
end if

dim Errfp
Errfp=false
dim ErrStr
ErrStr=""
dim show_old_pk
show_old_pk=oldmaxpk

if oldmaxpk=11 then show_old_pk="J"
if oldmaxpk=12 then show_old_pk="Q"
if oldmaxpk=13 then show_old_pk="K"
if oldmaxpk=14 then show_old_pk="A"
if oldmaxpk=15 then show_old_pk="2"
if oldmaxpk=16 then show_old_pk="�p��"
if oldmaxpk=17 then show_old_pk="�j��"

        IF old_pk_geshi=2 then
		if oldoldpn<3 then
			e_Str="�@��"
		else
			e_Str="�s��"
		end if
	elseif old_pk_geshi=3 then
		e_Str="�T��"
	elseif old_pk_geshi=4 then
		e_Str="���u"
	elseif old_pk_geshi=1 then
		if oldoldpn>4 then
			e_Str="�@���s�̤j���@�i�P�G"
		else
			e_Str="�@�i" 
		end if
	end if
	if (old_pk_geshi-5>0 and oldoldpn<>0) then e_Str=e_Str & "�M�@��" 

if zaid<>4 and oldoldpn<>0 then
        
	if old_pk_geshi<>pk_geshi then
  response.write "<script Language='Javascript'>alert('["& to1 & "��~�X���O[" & e_Str & " " & show_old_pk & "]��!');</script>"
  response.end
	elseif oldpn<>oldoldpn then
		'erralt(oldpn)
  response.write "<script Language='Javascript'>alert('["& to1 & "��~�X���O[" & oldoldpn & "]�i��!');</script>"
  response.end
	end if

	if ErrStr<>"" then
  response.write "<script Language='Javascript'>alert('�p�G�A�S���o�صP�A�Цb[�~�P]���[�� �P]!');</script>"
  response.end
	end if
end if


if pkyizha3=oldpn and oldpn>4 then 
qysall=qysall & "<b><font color=red>[�@���s]</font></b>"
IsyitiaoL=true
else
IsyitiaoL=false
end if
if zaid=0 and yidu=0 and sang=0 and IsyitiaoL=false then
isgood=false
else
isgood=true
end if


if (yidu>0 and pkyizha2>0) then Errfp=true 
if (zaid=0 and (yidu>2 and yidu<>liandui) ) then Errfp=true '���H�W���O�s��
if (sang<>0 and yidu=0 and pkyizha2>2) then Errfp=true '���T�Ӧ��@���٦��@�i�H�W�s�P

if (sang<>0 and oldpn<>5 and oldpn<>3) then Errfp=true '�T�Ӥj�󤭱i 
if (zaid<>0 and oldpn>7) then Errfp=true '���u�j��7�i

if (oldpn=3 and yidu<>0) then Errfp=true '�T�i�P�����@��

if (pkyizha2>2 and sang<>0) then Errfp=true '�p�G��i���O�@��
if (isgood=false and oldpn>1) then Errfp=true '�S�����S����S���T�Ӥ��O�@���s�A�ӥB���O�@�i�P
if Errfp=true then
  response.write "<script Language='Javascript'>alert('�A�Q�o���O�S���W�ߪ��s�P�A�o��j���H�����J�]���|�H�ݬ����U���I!');</script>"
  response.end
end if
if maxpk=14 then maxpk=16
if maxpk=15 then maxpk=17
'�s��14.15�����b��Ӫ�1415�᭱
if maxpk=1 then maxpk=14
if maxpk=2 then maxpk=15

'response.write maxpk
'response.end
if maxpk<=Cint(oldmaxpk) then
  if (zaid>0 and iszai) or zaid=0 then
  response.write "<script Language='Javascript'>alert('["& to1 & "��~�X���P[" & e_Str & " " & show_old_pk & "]��A�j�I�p�G�A�S���P���~���P�j�A�Цb[�~�P]���[�� �P]!');</script>"
  response.end
   end if
end if

if dpk3>0 then
sql="update dpk set pk='" & dpk & "',fp=true,oldpn=" & pk_geshi & oldpn & ",maxpk=" & maxpk & ",logtime='" & now() & "' where ufrom='"& info(0) & "'"
Set Rs=conn.Execute(sql)
sql="update dpk set fp=false,oldpn=" & pk_geshi & oldpn & ",maxpk=" & maxpk & ",iszai=" & iszai & " where ufrom='"& dpkto & "'"
Set Rs=conn.Execute(sql)
conn.close
dpkfp=info(0)&"�Q�F�Q�A���X[" & oldpn & "]�i�P�t�x���ϤW�F�P��A��ӬO......" & ismore & "<br> " & qysall & fpnow & " , �٦�" & dpk3 & "�i�P"
%>
<script language="JavaScript">parent.f3.location.reload();</script>
<%
elseif oldpn=18 then
sql="update dpk set fp=true,oldpn=90,maxpk=0 where ufrom='"& info(0) & "'"
Set Rs=conn.Execute(sql)
sql="update dpk set fp=false,oldpn=90,maxpk=0,iszai=false where ufrom='"& dpkto & "'"
Set Rs=conn.Execute(sql)
conn.close
  response.write "<script Language='Javascript'>alert('�A�Q�@����(����u�C���W�h�Q�ߥ��Ҧ����P)�H�{�b�n�@�A�A��[" & dpkto & "]�X�P!');</script>"
  response.end
else

Set Rs=conn.Execute(d_sql)
conn.close
	Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")

'------------------------�s�榡------------------------
sql="select " & xia_sql & " from �Τ� where �m�W='" & dpkto &"'"
set rs=conn.execute(sql)
yin1=rs(""&xia_sql&"")
if yin1<0 then yin1=0
if clng(yin1)<clng(xiazhu) then 
xiazhu=yin1
sayxia=",�o�a�먭�W�ӽa�F�A�u���o��h" & xia_class  & "�F"
end if

conn.execute "update �Τ� set " & xia_sql & "=" & xia_sql & "-"&xiazhu&" where �m�W='"& dpkto &"'"
conn.execute "update �Τ� set " & xia_sql & "=" & xia_sql & "+"& xiazhu&" where �m�W='"& info(0)&"'" 
conn.close
dpkfp=info(0)&"�����j���A��Ҧ����P�o�F�X��......" & ismore & "<br> " &fpnow&",�S���P�F,�q"&dpkto&"��Ĺ��" & xiazhu & xia_class & sayxia
'------------------------�s�榡------------------------

end if

elseif fn1="���n�F" then
sql="update dpk set fp=true,oldpn=90,maxpk=0 where ufrom='"& info(0) & "'"
Set Rs=conn.Execute(sql)
sql="update dpk set fp=false,oldpn=90,maxpk=0,iszai=false where ufrom='"& dpkto & "'"
Set Rs=conn.Execute(sql)
dpkfp=info(0)&"��󥴳o�i�P�A��[" & dpkto & "]�X�P"
conn.close
elseif fn1="�{��F" then

Set Rs=conn.Execute(d_sql)
conn.close
	Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")


'------------------------�s�榡------------------------
sql="select " & xia_sql & " from �Τ� where �m�W='" & info(0) &"'"
set rs=conn.execute(sql)
yin1=rs(""&xia_sql&"")
if clng(yin1)<clng(xiazhu) then 
xiazhu=yin1
sayxia="�A�o�a�먭�W�u���o��h" & xia_class  & "�F"
end if

conn.execute "update �Τ� set " & xia_sql & "=" & xia_sql & "+"&xiazhu&" where �m�W='"& dpkto &"'"
conn.execute "update �Τ� set " & xia_sql & "=" & xia_sql & "-"&xiazhu&" where �m�W='"& info(0) &"'"
conn.close
dpkfp=info(0)&"�{��F��󥴳o���P,�鵹[" & dpkto & "]" & xiazhu & xia_class & sayxia
'------------------------�s�榡------------------------

end if
set conn=nothing
set rs=nothing
end function

function showpk(pk)
dim wpk
wpk=replace(pk,"��","<img src=duju/dpk/1.GIF border=0><font color=#FF0000 size=2>")
wpk=replace(wpk,"��","<img src=duju/dpk/2.GIF border=0><font color=#000000 size=2>")
wpk=replace(wpk,"��","<img src=duju/dpk/3.GIF border=0><font color=#000000 size=2>")
wpk=replace(wpk,"��","<img src=duju/dpk/4.GIF border=0><font color=#FF0000 size=2>")
wpk=replace(wpk,"�p��","<img src=duju/dpk/xw.gif border=0>")
wpk=replace(wpk,"�j��","<img src=duju/dpk/dw.gif border=0>")
showpk=wpk & "</font>"
end function
function unconvpk(cpk)
dim pk2
pk2=replace(cpk,"A","1")
pk2=replace(pk2,"J","11")
pk2=replace(pk2,"Q","12")
pk2=replace(pk2,"K","13")
pk2=replace(pk2,"�p��","14")
pk2=replace(pk2,"�j��","15")
unconvpk=pk2
end function
%>
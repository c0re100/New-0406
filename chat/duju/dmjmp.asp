<%
function dmjGetmj(from1,to1)
if to1="�j�a" and fn1<>"�{��F" then
  response.write "<script Language='Javascript'>alert('�п�ܻ��ܹﹳ�A���๳�j�a�o�e�o�өR�O!');</script>"
  response.end
end if

chatroomsn=session("nowinroom")

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
db="duju/dmj.mdb"
'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(""&db&"")
'�p�G�A���A�Ⱦ����jet4.0�A�ШϥΤU�����s����k�A�����{�ǩʯ�
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.open connstr 

sql="select * from dmj where ufrom='"& info(0) & "'"
Set Rs=conn.Execute(sql)

if rs.eof and rs.bof then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('�z�S���ѥ[���±N�A���N�Q��N�P�F!');</script>"
  response.end
else
	isMy=rs("isMy")
	isFp=rs("isFp")
	isGet=rs("isGet")
	Mymj=rs("Mymj")
	dmjto=rs("uto")
	xiazhu=rs("duz")
	mjID=rs("mjID")
	logtime=rs("logtime")
	rs.close

if not isMy then
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('�{�b�n�����O�A�N�P��,�����X�P�F�~�O�A�N�P!');</script>"
  response.end
end if
if isFp and (not isGet) then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('�A�{�b�����P(�X�P),�b[�~�P]���I���±N�ϥ�(/���±N XXX) �R�O!');</script>"
  response.end
end if
sql="select �±N from mjInfo where id="& mjID 
Set Rs=conn.Execute(sql)
Getmjs=rs("�±N")
rs.close

if len(Getmjs)>39 then
Getmjs2=right(Getmjs,3)
Mymj=Mymj & Getmjs2
Getmjs=left(Getmjs,len(Getmjs)-3)
sql="update mjInfo set �±N='" & Getmjs & "' where id="& mjID 
Set Rs=conn.Execute(sql)
sql="update dmj set Mymj='" & Mymj & "',isGet=false,isFp=true,logtime='" & now() & "',mjmp='" & left(Getmjs2,2) & "' where ufrom='"& info(0) & "'"
Set Rs=conn.Execute(sql)
'last=strmj2(left(Getmjs2,2))

%>
<script language="JavaScript">parent.f3.location.href="duju/dmj-xp.asp";alert("�A�N��<%=strmj2(left(Getmjs2,2))%>")</script>
<%
dmjGetmj=info(0)&",�b�P��N�F�@�i�P�A�p�ߪ��ݤF�ݡA���F�����A�����b�Q����......"
else
sql="delete * from dmj where ufrom='" & info(0) & "' or ufrom='" & dmjto &"'"
Set Rs=conn.Execute(sql)
sql="delete * from mjInfo where id=" & mjID
Set Rs=conn.Execute(sql)
dmjGetmj=info(0)&",�u��13�i�P�F�A����A�N�P�A�A�̥����M���A�S���ӭt�C"
end if
end if
End Function
%>
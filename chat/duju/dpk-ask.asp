<%
function dpkask(fn1,from1,to1)

chatroomsn=session("nowinroom")

if to1="�j�a" or to1=info(0) then 
response.write "<script Language='Javascript'>alert('�A���i�H��ܤj�a�Φۤw�@�����C');</script>"
response.end
end if
'------------------------�s�榡------------------------
ARR_FN=Split(fn1,"|")
dim ub_Err,xia_class,xiamoney,xia_fir,xia_sql
if ubound(ARR_FN)<>1 Then 
response.write "<script Language='Javascript'>alert('�ާ@���~�G�R�O�榡�p�U�G\n /�����J �U�`�ƥ�|�Ȩ�[�ΡG�����A�k�O] \n\n[�ܨ�]�G\n /�����J 1000|�Ȩ�\n /�����J 1000|����\n  /�����J 1000|�k�O');</script>"
response.end
end if
xia_class=ARR_FN(1)
xiamoney=ARR_FN(0)

select case xia_class
	case "����"
		xia_fir="1"
		xia_sql="����"
	case "�Ȩ�"
		xia_fir="2"
		xia_sql="�Ȩ�"
	case "�k�O"
		xia_fir="3"
		xia_sql="�k�O"
	case else		
response.write "<script Language='Javascript'>alert('�ާ@���~�G�R�O�榡�p�U�G\n /�����J �U�`�ƥ�|�Ȩ�[�ΡG�����A�k�O] \n\n[�ܨ�]�G\n /�����J 1000|�Ȩ�\n /�����J 1000|����\n  /�����J 1000|�k�O');</script>"
response.end
end select

If not isnumeric(xiamoney) Then 
response.write "<script Language='Javascript'>alert('�ާ@���~�G�R�O�榡�p�U�G\n /�����J �U�`�ƥ�|�Ȩ�[�ΡG�����A�k�O] \n\n[�ܨ�]�G\n /�����J 1000|�Ȩ�\n /�����J 1000|����\n  /�����J 1000|�k�O');</script>"
response.end
end if


xiamoney=abs(int(xiamoney))

if (xia_fir="1")and(xiamoney<10 or xiamoney>1000) then 
response.write "<script Language='Javascript'>alert('�̤֩�G10" & xia_class & "�A�̦h1000�Ӫ����I');</script>"
response.end
end if
if (xia_fir="2")and(xiamoney<1000000 or xiamoney>10000000) then
response.write "<script Language='Javascript'>alert('�̤֩�G1000000" & xia_class & "�A�̦h1000�U�Ȩ�I');</script>"
response.end
end if
if (xia_fir="3")and(xiamoney<1000 or xiamoney>100000) then 
response.write "<script Language='Javascript'>alert('�̤֩�G1000" & xia_class & "�A�̦h10�U�k�O�I');</script>"
response.end
end if

'------------------------�s�榡------------------------

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")


'------------------------�s�榡------------------------
sql="select " & xia_sql & " from �Τ� where �m�W='" &info(0)&"'"
set rs=conn.execute(sql)
yin1=rs(""&xia_sql&"")

if xiamoney> yin1 then 
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
response.write "<script Language='Javascript'>alert('�A�����W�n���S���o��h" & xia_class&"');</script>"
response.end
end if
sql="select " & xia_sql & " from �Τ� where �m�W='" &to1&"'"
set rs=conn.execute(sql)
yin2=rs(""&xia_sql&"")
if xiamoney>yin2 then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
response.write "<script Language='Javascript'>alert('��設�W�n���S���o��h" & xia_class&"');</script>"
response.end
end if

'------------------------�s�榡------------------------
conn.close
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
db="duju/DPK.mdb"

'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(""&db&"")
'�p�G�A���A�Ⱦ����jet4.0�A�ШϥΤU�����s����k�A�����{�ǩʯ�
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.open connstr 
sql="select * from dpk where ufrom='"& info(0) & "'"

Set Rs=conn.Execute(sql)
if not (rs.eof and rs.bof) then
rs.close
conn.close
response.write "<script Language='Javascript'>alert('�z�٦��P���S�������A����}��');parent.f3.location.href='duju/dpk-xp.asp';</script>"
response.end
else
rs.close
'------------------------�s�榡------------------------
dpkask="<b><font color=green>["&info(0)&"]</font></b>��<b><font color=black>["&to1&"]</font></b>���G�����J�ܡH��`��<b><font color=red>"&xiamoney & xia_class & "</font></b>�A<img src='duju/dpk/1.GIF'><input type=button value='�@�N' onclick=javascript:window.open('duju/dpkok.asp?id="&myid&"&from1="&info(9)&"&to1="&to1&"&yn=1','a','width=380,height=210');this.disabled=1 style=background-color:#86A231;color:FFFFFF;border: 1 double onMouseOver=this.style.color='FFFF00' onMouseOut=this.style.color='FFFFFF' name=button223><img src='duju/dpk/2.GIF'><input type=button value='�ڵ�' onclick=javascript:window.open('duju/dpkok.asp?id="&myid&"&from1="&info(9) & "&to1="&to1&"&yn=0','a','width=380,height=210');this.disabled=1 style=background-color:#86A231;color:FFFFFF;border: 1 double onMouseOver=this.style.color='FFFF00' onMouseOut=this.style.color='FFFFFF' name=button223>"

sql="insert into dpk(ufrom,uto,duz) values ('"& info(0) & "$�U�`','"& to1 & "$�U�`', " & xia_fir & xiamoney & ")"
Set Rs=conn.Execute(sql)

conn.close
end if
set rs=nothing
set conn=nothing
end function
%>

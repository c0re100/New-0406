<%'�D�D
function titl(fn1)
if info(3)<10 then
	Response.Write "<script language=JavaScript>{alert('�߰ʦb�u�A�ݭn�԰�����10�šA�A�����r�I�I');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ���O,���� FROM �Τ� WHERE id=" & info(9),conn
if rs("���O")<1000 or rs("����")<10 then
	Response.Write "<script language=JavaScript>{alert('�ݭn���O1000�B����10�~�i�H�A�n�n�m�X�ѧa�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
else
	conn.execute "update �Τ� set ���O=���O-1000 where id=" & info(9)
	titl="<marquee border=1  scrolldelay=175><img src=dog.gif>" & fn1 & "�i"&info(0)&"�j</marquee>"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>
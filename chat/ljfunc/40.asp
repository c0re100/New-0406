<%'�߸�
function xintiao(fn1)
if info(3)<30 then
	Response.Write "<script language=JavaScript>{alert('���ާ@�ݭn�԰�����30�A�~�i�H�ާ@�I�I');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ���O FROM �Τ� WHERE id="&info(9),conn
if rs("���O")<1000 then
	Response.Write "<script language=JavaScript>{alert('�ݭn���O1000�~�i�H�ާ@�I�A�����r�I�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "update �Τ� set ���O=���O-1000 where id="&info(9)
xintiao="<marquee height=50 behavior=alternate loop=200 direction=up> �߸��߻y " & fn1 & "(" & info(0) & ")" & "</marquee>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
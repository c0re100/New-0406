<%
function xindong(fn1)
if info(3)<15 then
	Response.Write "<script language=JavaScript>{alert('�߰ʦb�u�A�ݭn�԰�����15�šA�A�����r�I�I');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ���O FROM �Τ� WHERE id="&info(9),conn
if rs("���O")<1000 then
	Response.Write "<script language=JavaScript>{alert('�߰ʡG�ݤ��O1000�A�A�����r�I�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
else
	conn.execute "update �Τ� set ���O=���O-1000 where id="&info(9)
	xindong="<marquee width=100% behavior=alternate scrollamount=10><img src=p71.gif>" & fn1 & "�i" & info(0) & "�j" & "</marquee>"
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
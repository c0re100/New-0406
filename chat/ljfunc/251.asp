<%'�U�s
function xiashan(to1,toname)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �v��,�Ȩ�,���� from �Τ� where id="&info(9),conn
sf=rs("�v��")
	if sf="" or sf="�L" then
		Response.Write "<script language=JavaScript>{alert('�A�S���v�šA�L�k�����v���I');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
if rs("�Ȩ�")<80000 then
	Response.Write "<script language=JavaScript>{alert('�A�S��8�U���A�H�a���z�r�I�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
	conn.execute "update �Τ� set �Ȩ�=�Ȩ�-50000,�v��='�L',�O�d1='�O�d' where id="&info(9)
	xiashan=info(0) & "�V��v��"& sf &"��ǤF8�U��������O�A�ש�P"& sf &"�����F�v�{��ô�I"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>

<%'���v
function bais(to1,toname)
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('���v�ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
if info(3)<2 then
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n2�ťH�W�~�i�H���v�I');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �v��,�Ȩ�,���� from �Τ� where id="&info(9),conn
sf=rs("�v��")
if sf=toname then
	Response.Write "<script language=JavaScript>{alert('["& toname&"]�w�g�O�A�v�ŤF');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End

end if
if rs("�v��")<>"" and rs("�v��")<>"�L" then
	Response.Write "<script language=JavaScript>{alert('�Q��["& toname&"]���v�A�лP�{�b�v�Ų�����ô�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("�Ȩ�")<50000 then
	Response.Write "<script language=JavaScript>{alert('�A�S��5�U��["& toname&"]���Q���A���̤l�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
dj=rs("����")
rs.close
rs.open "select ���� FROM �Τ� WHERE �m�W='"&toname&"'",conn
if rs("����")<30 or dj>rs("����") then
	Response.Write "<script language=JavaScript>{alert('["& toname&"]�٤���30�ŧr,���ব�̤l�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
bais=info(0) & "���ǳƦV"& toname &"��ǤF5�U�����v�O�A���X�F���v�ӽСA�]�����D�H�a�P�N���I"
Application("ljjh_bais_sf")=toname
Application("ljjh_bais_td")=info(0)
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>

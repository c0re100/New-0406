<%'�D�c
function qqq(to1,toname)
if toname="�j�a" then
	Response.Write "<script language=JavaScript>{alert('�D�c�ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �G�B,�t��,�ʧO,�Ȩ�,���� from �Τ� where id="&info(9),conn
sf=rs("�G�B")
po=rs("�t��")
xbk=rs("�ʧO")
if info(0)=toname then
	if sf="" or sf="�L" then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
if info(1)="�k" then
		Response.Write "<script language=JavaScript>{alert('�A�S���p�ѱC�A�L�k�𱼡I');}</script>"
else
Response.Write "<script language=JavaScript>{alert('�u�A�p�Ѥ��A�]�n��r�i�O�A�S���p�Ѥ��A�L�k�𱼡I');}</script>"
end if
		Response.End
	end if
if rs("�Ȩ�")<5000000 then
	Response.Write "<script language=JavaScript>{alert('�A�S��500�U���A�H�a���z�r�I�I');}</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end if
	conn.execute "update �Τ� set �Ȩ�=�Ȩ�-5000000,�G�B='�L' where id="&info(9)
        conn.execute "update �Τ� set �Ȩ�=�Ȩ�+5000000,�G�B='�L' where �m�W='"& sf &"'"
        conn.execute "delete * from �X��� where �m�W�k='" & info(9) & "' or �m�W�k='" & sf & "'"
       rs.close
	set rs=nothing
	conn.close
	set conn=nothing
if info(1)="�k" then
	qqq=info(0) & "����Ӫ��p�ѱC<font color=red>"& sf &"</font>��ǤF500�U��������O�A�ש��<font color=red>"& sf &"</font>����F~~~"
else
qqq=info(0) & "����Ӫ��p�Ѥ�<font color=red>"& sf &"</font>��ǤF1000�U��������O�A�ש��<font color=red>"& sf &"</font>���ϤF~~~"
end if
	exit function
end if
if sf=toname or po=toname then
	Response.Write "<script language=JavaScript>{alert('["& toname&"]�w�g�O�A���H�F�I');}</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
        Response.End

end if
if rs("�G�B")<>"" and rs("�G�B")<>"�L" then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing	
if info(1)="�k" then
Response.Write "<script language=JavaScript>{alert('�Q�M["& toname&"]�n�A�лP�{�b���p�ѱC������ô�I');}</script>"
else
Response.Write "<script language=JavaScript>{alert('�Q�M["& toname&"]�n�A�лP�{�b���p�Ѥ�������ô�I');}</script>"
end if
	Response.End
end if
if rs("�Ȩ�")<5000000 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing	
Response.Write "<script language=JavaScript>{alert('�A�S��500�U��["& toname&"]���Q�M�A�n�I');}</script>"
	
	Response.End
end if
rs.close
rs.open "select �ʧO,�t�� FROM �Τ� WHERE �m�W='"&toname&"'",conn
if rs("�ʧO")=xbk  then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing	
Response.Write "<script language=JavaScript>{alert('�d����ڡH���򤣤��\�s�b�P���ʪ��I');}</script>"
	Response.End
end if

if info(1)="�k" then
qqq=info(0) & "�V"& toname &"��F500�U��§���A���X�F�n����謰�p�ѱC�A�]�����D�H�a�P�N���I"
else
qqq=info(0) & "�V"& toname &"��F1000�U��§���A���X�F�n����谵�ۤv���p�Ѥ��A�]�����D�H�a�P�N���I"
end if
Application("ljjh_qqq_sf")=toname
Application("ljjh_qqq_td")=info(0)
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>

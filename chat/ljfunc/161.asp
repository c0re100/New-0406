<%'���W�B�m
function mengui(fn1,to1,toname)
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('�B�m�ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ����,���� FROM �Τ� WHERE id="&info(9),conn
fn1=int(abs(fn1))
mpai1=rs("����")
if rs("����")<>"�x��" and rs("����")<>"����" and rs("����")<>"�ƴx��" then
	Response.Write "<script language=JavaScript>{alert('�A�O���򨭥��H�n���ѯťH�W����~�����\��I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if fn1>50000 then
	Response.Write "<script language=JavaScript>{alert('�@�ڳ̦h�u��5�U�H�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
rs.close
rs.open "select ����,�Ȩ�,�s�� FROM �Τ� WHERE �m�W='"&toname&"'",conn
mpai2=rs("����")
if mpai1<>mpai2 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
  Response.Write "<script language=JavaScript>{alert('��褣�O�A�������̤l�ڡH�I');}</script>"
Response.End
end if
if rs("�Ȩ�")>=fn1 or rs("�s��")>=fn1 then
if rs("�Ȩ�")>=fn1 then
conn.execute "update ���� set �������=�������+"& fn1 &" where ����='" &  info(5) &"'"
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & fn1 & " where �m�W='"&toname&"'"
mengui="�x���G"&info(0) & "�]�̤l" & toname & "�H�Ϫ��W�A�����W�B�m�@��" & fn1 &"��,�һ@�ڶ���J�������......"
end if
if rs("�s��")>=fn1 then
conn.execute "update ���� set �������=�������+"& fn1 &" where ����='" &  info(5) &"'"
conn.execute "update �Τ� set �s��=�s��-" & fn1 & " where �m�W='"&toname&"'"
mengui="�x���G"&info(0)& "�]�̤l" & toname & "�H�Ϫ��W�A�����W�B�m�@��" & fn1 &"��,�һ@�ڶ���J�������......"
end if
else
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
  Response.Write "<script language=JavaScript>{alert('�A���̤l���W�S�X��Ȥl�ڡH�I');}</script>"
Response.End
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>

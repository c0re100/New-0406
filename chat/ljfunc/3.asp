<%'�ѥ�
function jie(fn1)
fn1=trim(fn1)
if info(2)<6 then
	Response.Write "<script language=JavaScript>{alert('�A���O�x�����H�A����ѥޡI');}</script>"
	Response.End
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select ���A,�n�� from �Τ� where �m�W='" & fn1 & "'",conn
if trim(rs("���A"))="�I��" and (not rs.eof) and rs("�n��")>now() then
	conn.execute "update �Τ� set �n��=now(),���A='���`' where �m�W='" & fn1 & "'"
	jie=info(0)& "��" & fn1 & "�ϥΤF�ѥ޳N�A" & fn1 & "�r�M���L�ӤF  "
else
	jie=info(0)& "��" & fn1 & "�ϥΤF�ѥ޳N�A�i�O" & fn1 & "�èS���Q�I��  "
end if
conn.execute "insert into �ާ@�O��(�ɶ�,�m�W1,�m�W2,�覡,�ƾ�) values (now(),'"& info(0) &"','"& fn1 &"','�ѥ�','123')"
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end function
%>

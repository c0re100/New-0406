<%'�٭��N
function dddd(to1,fn1,toname)
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('�ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
if info(2)<20000000 then
	Response.Write "<script language=JavaScript>{alert('�Q�@����r�A20000000�ũx���~�i�Ϊ��I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select grade,�m�W from �Τ� where �m�W='"&toname&"'",conn
toname=rs("�m�W")
if rs.eof or rs.bof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�S���o�ӤH�H�A�O���O�ݿ��F�I');}</script>"
	
	Response.End
end if
if rs("grade")<info(2)  then
	call boot(toname)
	conn.execute "update �Τ� set ���A='��' where �m�W='"&toname&"'"
	dddd=info(0)& "��" & toname & "�ϥΤF�٭��N�A" & toname & "�ߨ覺�` "
else
	dddd=info(0) & "��" & toname & "�ϥΤF�٭��N�A�i�O�L�O�x�������ź޲z���I"
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function
%>

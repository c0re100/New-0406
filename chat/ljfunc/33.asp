<%'���
function zzyin(fn1,to1,toname)
fn1=abs(fn1)
if fn1<10000 or fn1>5000000 then
	Response.Write "<script language=JavaScript>{alert('���̤�1�U�A�̦h500�U�I');}</script>"
	Response.End
end if
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('���ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
if info(3)<2 then
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n2�ťH�W�~�i�H���I');}</script>"
	Response.End
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select �s�� from �Τ� where id="&info(9),conn
money=rs("�s��")
if rs("�s��")<fn1 then
	Response.Write "<script language=JavaScript>{alert('�A������h���s�ڶܡH�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
moneyok=money+fn1
if moneyok>2000000000 then
	Response.Write "<script language=JavaScript>{alert('��誺�s�ڤӦh�N�ֶW�L�F20���F�A���������{���бz�������ǧa�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if

conn.execute "update �Τ� set �s��=�s��-"& fn1 &" where id="&info(9)
conn.execute "update �Τ� set �s��=�s��+"& int(fn1*0.9) &" where �m�W='"&toname&"'"
zzyin=info(0) & "��ۤv���򪺦s��:"& fn1 &"��A����"& toname &"���Ȧ�W�U�A"&toname&"�V�x�������O5%�����ާ@���\�I"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
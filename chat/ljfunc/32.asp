<%'����
function getyin(fn1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select �s��,������,�Ȩ�,�ާ@�ɶ� from �Τ� where id="&info(9),conn
if DateDiff("s",rs("�ާ@�ɶ�"),now())<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���F����b�Ȧ����A�s������10�����ާ@!');}</script>"
	Response.End 
end if
bankmoney=rs("�s��")
lastdate=date()-rs("������")
money=rs("�Ȩ�")
moneyok=int(money)+abs(fn1)
if moneyok>2000000000 then
	Response.Write "<script language=JavaScript>{alert('�A���{���Ӧh�N�ֶW�L�F20���F�A���������{���бz����s�Ƕi�ѥ��a�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
newbankmoney=int(bankmoney+bankmoney*0.0001*lastdate)
fn1=abs(fn1)
if fn1<=0 then
	if bankmoney=<0 then
		getyin=info(0) & "�A�b��������èS���s��,�����C�ѧQ����:0.0001%,�w��s��!"
	else
		getyin=info(0) & "�b�����s��:"& bankmoney &"��,�b:"& rs("������") &"�s�J,��0.0001%�Q,�Ȧ�{�b��:"& newbankmoney &"��!"
	end if
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if newbankmoney<fn1 then
	Response.Write "<script language=JavaScript>{alert('�A���̦�����h�����H�H');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
getyin=info(0) & "�V��������:"& fn1 &"��,�����{���Ȩ�:"& newbankmoney-fn1 &"��,���n�A����,�O�Q�m�F!"
conn.execute "update �Τ� set �Ȩ�=�Ȩ�+"  & fn1 & ",�s��="& newbankmoney-fn1 &",������=date(),�ާ@�ɶ�=now() where id="&info(9)
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
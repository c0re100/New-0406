<%'���k�O
function getfali(fn1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select �k�O�I,������,�k�O,�ާ@�ɶ� from �Τ� where id="&info(9),conn
if DateDiff("s",rs("�ާ@�ɶ�"),now())<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���F����b�k�O�_�~�è�A�s���k�O�I��10�����ާ@!');}</script>"
	Response.End 
end if
bankmoney=rs("�k�O�I")
lastdate=date()-rs("������")
money=rs("�k�O")
newbankmoney=int(bankmoney+bankmoney*0.0001*lastdate)
fn1=abs(fn1)
if fn1<=0 then
	if bankmoney=<0 then
		getyin=info(0) & "�A�b�k�O�_�~�èS���s��k�O�I,�x���N���O�ުk�O�_�~�C�ѧQ����:0.0001%,�w��s��!"
	else
		getyin=info(0) & "�b�k�O�_�~�̦@�s��:"& bankmoney &"�I�k�O,�b:"& rs("������") &"�s�J,��0.0001%�Q,�k�O�_�~�̲{�b��:"& newbankmoney &"�I�k�O!"
	end if
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if newbankmoney<fn1 then
	Response.Write "<script language=JavaScript>{alert('�A���̦�����h���k�O�ڡH�H');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
getfali=info(0) & "�q�k�O�_�~�̨��X:"& fn1 &"�I�k�O,�k�O�_�~�{�s���k�O:"& newbankmoney-fn1 &"�I,�n�n�ϥ�,�O�Q�O�H�l�F!"
conn.execute "update �Τ� set �k�O=�k�O+"  & fn1 & ",�k�O�I="& newbankmoney-fn1 &",������=date(),�ާ@�ɶ�=now() where id="&info(9)
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
<%'�e��
function daipu(fn1,to1,toname)
if info(2)<6 then
	Response.Write "<script language=JavaScript>{alert('�A�L�v�ާ@�A�A���O�x�����H�I');}</script>"
	Response.End
end if
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('�e���ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select �m�W,grade,���� from �Τ� where �m�W='"&toname&"'",conn
toname=rs("�m�W")
if info(2)<=rs("grade") then
	Response.Write "<script language=JavaScript>{alert('���ѡA�A����ﰪ�ź޲z���ާ@�θӤH�O�S�O�O�@�ﹳ!');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
rs.close
rs.open "select id FROM �Τ� WHERE �m�W='"&toname&"'",conn
idd=rs("id")
rs.close
rs.open "select ���~�W FROM ���~ WHERE ����='�d��'  and �ƶq>0 and ���~�W='�j�K�d' and �֦���="& idd,conn
if not (rs.eof or rs.bof)  then

conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&idd&" and ���~�W='�j�K�d'")
	Response.Write "<script language=JavaScript>{alert('��設�W�{���j�K�d�A�е��L�����|���n��L�j�A�������ӭ��lOK�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "update �Τ� set ���A='��',�n��=now()+3 where �m�W='"&toname&"'"
call boot(toname)
conn.execute "insert into �ư�(�޲z����,�޲z�ﹳ,�޲z�覡,�欰�ɶ�,��]) values ('" & info(0) & "','" & toname & "','�e��',now(),'" & fn1 & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
if info(5)<>"�x��" then
daipu=info(0) & "[�K����]:<font color=red><bgsound src=wav/oh_no.wav loop=1>�x�����H���F���K���M�b" & toname & "����l�W,�s�Աa�誺��" & toname & "���F�c�СI[" & fn1 & "]</font>"
else
daipu=info(0) & "[�x���H��]:<font color=red><bgsound src=wav/oh_no.wav loop=1>�x�����H���F���K���M�b" & toname & "����l�W,�s�Աa�誺��" & toname & "���F�c�СI[" & fn1 & "]</font>"
end if
end function
%>
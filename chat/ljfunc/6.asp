<%'�ʸT
function jianjin(fn1,to1,toname)
if info(2)<8 then
	Response.Write "<script language=JavaScript>{alert('�ʸT�ݭn8�ťH�W�޲z���ާ@�I');}</script>"
	Response.End
end if
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('�ʸT�ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select id,grade from �Τ� where �m�W='"&toname&"'",conn
idd=rs("id")
if info(2)<=rs("grade")  then
	Response.Write "<script language=JavaScript>{alert('���i�H�ﰪ�ź޲z���ާ@�θӤH�O�S�O�O�@�ﹳ�I');}</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end if
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
conn.execute "update �Τ� set ���A='�ʸT',�n��=now()+1 where �m�W='"&toname&"'"
call boot(toname)
if info(5)<>"�x��" then
jianjin= info(0) &"[�K����]:<font color=red><bgsound src=wav/daipu.wav loop=1>�x�����H���F���K���M�b" & toname & "����l�W,�@�}��" & toname & "��i�F�c��,�ʵۧa,�z�w�g�Q�P�B�ר��ʸT!   [" & fn1 & "]</font>"
else
jianjin= info(0) &"[�x���H��]:<font color=red><bgsound src=wav/daipu.wav loop=1>�x�����H���F���K���M�b" & toname & "����l�W,�@�}��" & toname & "��i�F�c��,�ʵۧa,�z�w�g�Q�P�B�ר��ʸT!   [" & fn1 & "]</font>"
end if
conn.execute "insert into �ư�(�޲z����,�޲z�ﹳ,�޲z�覡,�欰�ɶ�,��]) values ('" & info(0) & "','" & toname & "','�ʸT',now(),'" & fn1 & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
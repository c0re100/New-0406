<%'���Y
function touzi(fn1,to1,toname)
if toname="" or toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('���Y�ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
'�P�_���H��
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ���H��,�O�@,����,����,��O,���O from �Τ� where id="&info(9),conn
if rs("���H��")>=30 then
	
	Response.Write "<script language=JavaScript>{alert('�A�@�c�Ӧh�A�W�L������H�W��30����A���Y�F�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if

if rs("��O")<300 or rs("���O")<300  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=javascript>alert('�i"&info(0)&"�j�A�{�b�����O����O�p��300�Ф��n�i����Y�R�O�ާ@�I');</script>"
	Response.End
end if
if rs("�O�@")=true then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�Ш����m�\�O�@�A�ާ@�I');</script>"
	Response.End
end if
if rs("����")<2 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n2�ťH�W�~�i�H�l�O�H���O�A���M�N�|��¼¼���I');}</script>"
	Response.End
end if
if rs("����")="�X�a" then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('�A�w�g�X�a�F�A���i�H�A���H�F�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
rs.open "select grade,����,�O�@,��O from �Τ� where �m�W='"&toname&"'",conn
if rs("����")<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=javascript>alert('�i"&toname&"�j�O�s��n�`�N�O�@�A������Y�F��I');</script>"
Response.End
	exit function
end if
if rs("�O�@")=true then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��西�b�����I');</script>"
	Response.End
end if
if rs("��O")<0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���w�g���F�I');</script>"
	Response.End
end if
rs.close
'�}�l���~
rs.open "select id,���O,��O,�ƶq from ���~ where ����='�t��' and �ƶq>0 and ���~�W='" & fn1 & "' and �֦���="&info(9) ,conn
if rs.eof or rs.bof  then
	Response.Write "<script language=JavaScript>{alert('�A���̦�["& fn1 &"]�o�˪��~�r�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
id=rs("ID")
neili=abs(rs("���O"))
tili=abs(rs("��O"))
'���~-1�ìd��0�ɧR���I
conn.execute "update ���~ set �ƶq=�ƶq-1 where id="&id
'�}�l���
conn.execute "update �Τ� set �D�w=�D�w-20,��O=��O-" & int(neili/4) & " where id="&info(9)
conn.execute "update �Τ� set ���O=���O-" & neili & ",��O=��O-" & tili & " where �m�W='"&toname&"'"
rs.close
rs.open "select �m�W,��O from �Τ� where �m�W='"&toname&"'",conn
toname=rs("�m�W")
if rs("��O")>-100 then
	touzi=info(0) & "<bgsound src=wav/anqi.wav loop=1>�V" & toname & "�U" & fn1 & ",��" & toname & "�����O��" & neili & "��O��" & tili & "!�ۤv�o���Ӥ��O:-"& int(neili/4 ) &"�I!"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if
conn.execute "update �Τ� set ���H��=���H��+1 where id="&info(9)
touzi=info(0) & "<bgsound src=wav/daipu.wav loop=1>�V" & toname & "���Y�F" & fn1 & ",��" & toname & "�����O:-" & neili & "��O:-" & tili & "," & toname & "�|��e���D���d�X����,���ڳ���  ��"
conn.execute "update �Τ� set ���A='��' where �m�W='"&toname&"'"
'�O��
conn.execute "insert into �H�R(����,�ɶ�,����,���]) values ('" & toname & "',now(),'" & info(0) & "','"&fn1&"')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
call boot(toname)
end function
%>
<%'�аV�Ĥl
function hzb(to1,toname)
if info(11)="���L" then
	hzb=info(0)&",���S�d���r�A�A���Ĥl�N���F��A�A�ٱаV�H�a�I"
	exit function
end if 
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('�аV�Ĥl�ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT ����,�p�� FROM �Τ� WHERE  id="&info(9),conn
fromdj=rs("����")
if rs("�p��")="���L" then
	hzb=info(0)&",���S�d���r�A�A���Ĥl�N���F��A�A�ٱаV�H�a�I"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if 
rs.close
rs.open "SELECT id,����,�p��,�ĦW FROM �Τ� WHERE �m�W='"&toname&"'",conn
idd=rs("id")
to1dj=rs("����")
to1dd=rs("�ĦW")
if rs("�p��")="�L" then
	hzb=info(0)&",�A�o���򯫸g,"&toname&"�٨S�Ĥl�O�I"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if 
if rs("�p��")="��" then
	hzb=info(0)&","&toname&"���Ĥl["&to1dd&"]�S���L�F��,�A����ðV!"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if 

if rs("�p��")="���L" then
randomize timer
r=int(rnd*7)+1
Select Case r
Case 1
	conn.execute "update �Τ� set �p��='��' where �m�W='"&toname&"'"
	conn.execute "update �Τ� set �D�w=�D�w+100  where id="&info(9)
	hzb="�����I"&toname&"���Ĥl["&to1dd&"]�Q"&info(0)&"�аV�F�@�y�A��H�{����!("&info(0)&"�D�w����100�I)"

Case 2
	if todj>fromdj then
		conn.execute "update �Τ� set �y�O=�y�O-100  where id="&info(9)
		conn.execute "update �Τ� set �p��='��' where �m�W='"&toname&"'"
		hzb="�K�K�I"&toname&"���Ĥl["&to1dd&"]���`�ӱ�,"&info(0)&"�ڥ����O�Ĥl�����,�ש����p��"&toname&"���]�F("&info(0)&"�y�O�U��100�I)"
	else
		hzb=info(0)&"���Q�e��"&toname&"���Ĥl["&to1dd&"],�J�Ӥ@�@,��ӬO�j���l,��O��"&toname&"���īĤl��F!"
	end if 
Case 3
	conn.execute "update �Τ� set �y�O=�y�O+100,�D�w=�D�w+100,�Ȩ�=�Ȩ�+2000000,���O=���O-200  where id="&info(9)
	conn.execute "update �Τ� set ���A='��',�n��=now()+3,�p��='��',���O=���O/2,��O=��O/2 where �m�W='"&toname&"'"
	call boot(toname)
	hzb="�z�I,"&info(0)&"�N"&toname&"���Ĥl["&to1dd&"]�M�j�H�@�_�j�|�F�@�y�I"&toname&"ı�o�S���l��a�F�A����Z�Ҥ����X�ӤF�C("&info(0)&"��o����2000000��,�D�w����100�I,�y�O����100�I,���O�U��200�I�C"&toname&"��O�M���O�U���@�b)"

Case 4
	conn.execute "update �Τ� set �p��='��' where �m�W='"&toname&"'"
	conn.execute "update �Τ� set �y�O=�y�O-100  where id="&info(9)
rs.close
	rs.open "select id,���~�W,����,���O,��O,����,���s,��T��,����,�Ȩ�,���� from ���~ where  �ƶq>0 and �˳�=false and  ����<>'�d��' and �֦���="&idd ,conn
	id=rs("id")
	wupin=rs("���~�W")
	lx=rs("����")
	nl=rs("���O")
	tl=rs("��O")
	gj=rs("����")
	fy=rs("���s")
	yin=rs("�Ȩ�")
        jgd=rs("��T��")
	dj=rs("����")
	if lx="�A��" or lx="�k��" or lx="����" or lx="���~" or lx="�k��" or lx="�k�_" or lx="�j��" or lx="�u��" or lx="�ī~" or lx="���}" or lx="�Y��" or lx="�˹�" or lx="����" or lx="�t��" then
		sm=rs("����")
	else
		sm=rs("����")
	end if
	rs.close
    conn.execute "update ���~ set �ƶq=�ƶq-1 where id="&id
	rs.open "select id from ���~  where �˳�=false and ���~�W='" & wupin & "' and �֦���=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~ (���~�W,�֦���,����,��T��,����,���s,���O,��O,�ƶq,�Ȩ�,����,����) values ('"&wupin&"','"&info(0)&"','"&lx&"',"&jgd&","&gj&","&fy&","&nl&","&tl&",1,"&yin&",'"& sm &"',"& dj &")"
	else
		id=rs("id")
		conn.execute "update ���~ set �Ȩ�=" & yin & ",�ƶq=�ƶq+1,����='"& sm &"' where id="&id
	end if
	hzb="�H�I"&toname&"���Ĥl["&to1dd&"]�D�`�F��,���פF"&info(0)&"�A�ܬӦӰk,���W�����~["&wupin&"]���F�U��,�Q�H�a�ߨ�(�y�O�U��100�I)"

case else
	hzb=info(0)&"���Q�аV"&toname&"���Ĥl["&to1dd&"],�J�Ӥ@�@,��ӬO�j���l,��O��"&toname&"���Ĥl����F!"
End Select

end if 
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function 
%>
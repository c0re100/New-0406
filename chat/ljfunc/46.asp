<%'��p��
function pk(to1,toname)
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('��p���ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
if info(4)=0 then 
aaao=0
else
aaao=1
end if
if info(3)<2 then
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n2�ťH�W�~�i�H��p���I');}</script>"
	Response.End
end if
if left(info(8),2)="�p��" then
	pk=info(0)&",���S�d���r�A�A�N�O�p���A�٧줰��p���r�I"
	exit function
end if 
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT ���� FROM �Τ� WHERE  id="&info(9),conn
fromdj=rs("����")
rs.close
rs.open "SELECT ����,����,¾�~ FROM �Τ� WHERE  �m�W='"&toname&"'",conn
to1dj=rs("����")
if rs("����")="�x��" then
   pk=info(0)&",�A���S���d���r�A�x�����|���p��!"
   rs.close
   set rs=nothing
   conn.close
   set conn=nothing
   exit function
end if

if rs("¾�~")<>"�p��(�b�k)" then
	pk=info(0)&",�A�o���򯫸g,"&toname&"�i�S���L�F��,�A��L������!"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if 

if rs("¾�~")="�p��(�b�k)" then
randomize timer
r=int(rnd*7)+1
Select Case r
Case 1
	conn.execute "update �Τ� set ¾�~='�Ԥh' where �m�W='"&toname&"'"
	conn.execute "update �Τ� set �D�w=�D�w+200  where id="&info(9)
	pk="�p��"&toname&"�b"&info(0)&"���оɤU,�M�w��צۭ�,��L�۷s���A���p��!("&info(0)&"�D�w����200�I)"

Case 2
	if todj>fromdj then
		conn.execute "update �Τ� set �y�O=�y�O-100  where id="&info(9)
		conn.execute "update �Τ� set ¾�~='�Ԥh' where �m�W='"&toname&"'"
		pk="�p��"&toname&"�Z�\���j,"&info(0)&"�ڥ����O�p�������,�ש����p��"&toname&"���]�F("&info(0)&"�y�O�U��100�I)"
	else
		pk=info(0)&"���Q�e��p��"&toname&",�J�Ӥ@�@,��ӬO�ۮa�˱�,��O��p��"&toname&"��F!"
	end if 
Case 3
	conn.execute "update �Τ� set �y�O=�y�O+100,�D�w=�D�w+100,�Ȩ�=�Ȩ�+10000,���O=���O-200  where id="&info(9)
	conn.execute "update �Τ� set ¾�~='�Ԥh',���A='��' where �m�W='"&toname&"'"
	call boot(toname)
	pk="�g�L�@�f�W��,����"&info(0)&"�ש�N�p��"&toname&"�e���k��!(������o����10000��,�D�w����100�I,�y�O����100�I,���O�U��200�I)"

Case 4
	conn.execute "update �Τ� set ¾�~='�Ԥh' where �m�W='"&toname&"'"
	conn.execute "update �Τ� set �y�O=�y�O-100  where id="&info(9)
	rs.close
rs.open "select id from �Τ� where  �m�W='"&toname&"'",conn
idd=rs("id")
rs.close
rs.open "select id,���~�W,����,���O,��O,����,���s,�Ȩ�,����,��T��,���� from ���~ where  �ƶq>0 and �˳�=false and  ����<>'�d��' and �֦���="&idd ,conn
	id=rs("id")
	wupin=rs("���~�W")
	lx=rs("����")
	nl=rs("���O")
	tl=rs("��O")
	gj=rs("����")
	fy=rs("���s")
	yin=rs("�Ȩ�")
	dj=rs("����")
	jgd=rs("��T��")
	if lx="�A��" or lx="�k��" or lx="����" or lx="���~" or lx="�k��" or lx="�k�_" or lx="�j��" or lx="�u��" or lx="�ī~" or lx="���}" or lx="�Y��" or lx="�˹�" or lx="����" or lx="�t��" or lx="�r��" then
	sm=rs("����")
else
	sm=rs("����")
end if
	
    conn.execute "update ���~ set �ƶq=�ƶq-1,�|��="&aaao&" where id="&id
rs.close
	rs.open "select id from ���~  where �˳�=false and ���~�W='" & wupin & "' and �֦���=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~ (���~�W,�֦���,����,����,���s,���O,��O,�ƶq,�Ȩ�,����,����,��T��,�|��) values ('"&wupin&"','"&info(0)&"','"&lx&"',"&gj&","&fy&","&nl&","&tl&",1,"&yin&",'"& sm &"',"&dj&","&jgd&","&aaao&")"
	else
		id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="&id
	end if
	pk="�p��"&toname&"�D�`�F��,���פF����"&info(0)&"���l��,�ܬӦӰk,���W�����~["&wupin&"]���F�U��,�Q���־ߨ�(���־y�O�U��100�I)"

case else
	pk=info(0)&"���Q�e��p��"&toname&",�J�Ӥ@�@,��ӬO�ۮa�˱�,��O��p��"&toname&"��F!"
End Select

end if 
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function 
%>
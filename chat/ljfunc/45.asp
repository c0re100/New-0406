<%'�p��
function tdx(to1,toname)
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('���F��ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
if info(4)=0 then 
aaao=0
else
aaao=1
end if
if info(3)<3 then
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n3�ťH�W�~�i�H���F��I');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT ����,��O,���O,�ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
dj=rs("����")
if rs("��O")<300 or rs("���O")<300  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=javascript>alert('�i"&info(0)&"�j�A�{�b�����O����O�p��300�Ф��n�i�氽�F��I');</script>"
	Response.End
end if
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if sj<8 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=8-sj
	Response.Write "<script language=JavaScript>{alert('�ЧA���W["& ss &"]��,�A�ާ@�I');}</script>"
	Response.End
end if
conn.execute "update �Τ� set �D�w=�D�w-100,¾�~='�p��(�b�k)',�ާ@�ɶ�=now() WHERE  id="&info(9)

lxlx="("
if dj>=5 then 
	lxlx=lxlx&"����='�ī~'"
end if 
if dj>=6 then 
	lxlx=lxlx&" or ����='�r��'"
end if 
if dj>=7 then
	lxlx=lxlx&" or ����='�A��'"
end if 
if dj>=8 then
	lxlx=lxlx&" or ����='�t��'"
end if 
if dj>=9 then 
	lxlx=lxlx&" or ����='�˹�'"
end if 
if dj>=10 then 
	lxlx=lxlx&" or ����='����' or ����='�k��'"
end if 
if dj>=11 then 
	lxlx=lxlx&" or ����='�Y��' or ����='����' or ����='���}'"
end if 
lxlx=lxlx&")"
rs.close
rs.open "select id,grade from �Τ� where �m�W='"&toname&"'",conn
idd=rs("id")
rs.close
rs.open "select id,���~�W,����,���O,��O,����,���s,����,��T��,�Ȩ�,���� from ���~ where "&lxlx&" and �ƶq>0  and ����<>'�d��' and �˳�=false and �֦���="&idd,conn
if rs.eof or rs.bof then
            tdx=info(0)&"�A�A�i�u�O�˾`�r�A"&toname&"���W����]�S���A�A�]�Q���I"
            rs.close
            set rs=nothing
            conn.close
            set conn=nothing
            exit function
else
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
	conn.execute "update ���~ set �ƶq=�ƶq-1 where id="&id
	rs.close
	rs.open "select id from ���~  where �˳�=false and ���~�W='" & wupin & "' and �֦���="& info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~ (���~�W,�֦���,����,����,���s,���O,��O,�ƶq,�Ȩ�,����,����,��T��,�|��) values ('"&wupin&"',"&info(9)&",'"&lx&"',"&gj&","&fy&","&nl&","&tl&",1,"&yin&",'"&sm&"',"&dj&","&jgd&","&aaao&")"
	else
		id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="&id
	end if
	tdx=info(0)&"<bgsound src=wav/xiaotou.wav loop=1>�q"&toname&"�����W�Ѩ�["&wupin&"]1�ӡA�����o��R���СI<font color=red>(���D�w100�I)</font>"
end if     
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function 
%>
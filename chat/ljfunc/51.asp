<%'�Ĥl����
function hza(to1,toname)
if info(11)="�L" or info(13)="�L" then
	hza=info(0)&"�A�٨S���Ĥl�O�I"
	exit function
end if 
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('�Ĥl���ѹﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
if info(3)<2 then
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n2�ťH�W�~�i�H�Ĥl���ѡI');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT ����,�p��,�ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
dj=rs("����")
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
if rs("�p��")="�L" then
	hza=info(0)&"�S�Ĥl��򰽪F��H"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if 
conn.execute "update �Τ� set �D�w=�D�w-10,�Ȩ�=�Ȩ�-5000,�p��='���L',�ާ@�ɶ�=now() WHERE  id="&info(9)

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
	lxlx=lxlx&" or ����='�Y��' or ����='����' or ����='���}' or lx='���~' or lx='�k��' or lx='�k�_' or lx='�j��' or lx='�u��'"
end if 
lxlx=lxlx&")"
rs.close
rs.open "select id from �Τ� where �m�W='"&toname&"'",conn
idd=rs("id")
rs.close
rs.open "select id,���~�W,����,���O,��O,����,���s,�Ȩ�,����,��T��,���� from ���~ where "&lxlx&" and �ƶq>0  and ����<>'�d��' and �˳�=false and �֦���="&idd,conn
if rs.eof or rs.bof then
            hza=info(0)&"�A�A�i�u�O�˾`�r�A"&toname&"���W����]�S���A�A���Ĥl��򰽡H"
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
        jgd=rs("��T��")
	dj=rs("����")
	if lx="�A��" or lx="�k��" or lx="����" or lx="���~" or lx="�k��" or lx="�k�_" or lx="�j��" or lx="�u��" or lx="�ī~" or lx="���}" or lx="�Y��" or lx="�˹�" or lx="����" or lx="�t��" then
		sm=rs("����")
	else
		sm=rs("����")
	end if
	conn.execute "update ���~ set �ƶq=�ƶq-1 where id="&id
	rs.close
	rs.open "select ���~�W,id from ���~  where �˳�=false and ���~�W='" & wupin & "' and �֦���="& info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~ (���~�W,�֦���,����,��T��,����,���s,���O,��O,�ƶq,�Ȩ�,����,����) values ('"&wupin&"',"&info(9)&",'"&lx&"',"&jgd&","&gj&","&fy&","&nl&","&tl&",1,"&yin&",'"&sm&"',"&dj&")"
	else
		id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,����='"&sm&"' where id="&id
	end if
	hza=info(0)&"���Ĥl["&info(13)&"]<bgsound src=wav/xiaotou.wav loop=1>�q"&toname&"�����W�Ѩ�["&wupin&"]1�ӡA�@�L�֪����Ӽ�~~~<font color=red>(���D�w10�I�A�Ȩ�U��5000)</font>"
end if     
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function 
%>
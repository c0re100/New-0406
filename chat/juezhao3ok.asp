<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI"
location.href = "javascript:history.go(-1)"
</script>
<%end if%>
<%ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
to1=request.form("to1")
gw=left(to1,1)
if IsNumeric(gw)=false and Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(to1)&" ")=0 then
zz=split(to1,"��")
gw=zz(0)
	Response.Write "<script Language=Javascript>alert('�ҧ������H���b��ѫǡI');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if Application("jingda")=1 then
	Response.Write "<script Language=Javascript>alert('�ثe�ѩ�H�֪���]�t�γQ�]�m���T�����A�A�n�Q���[�Ч�x���H���}�ҡI');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
stunt=""
stunt1=""
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select top 1 �m�W1,�m�W2 FROM ��D ",conn
xx1=rs("�m�W1")
xx2=rs("�m�W2")
rs.close
rs.open "select ����,���A,�Z�\,����,�Ȩ�,���s,��O,���H��,�O�@,��O,�ާ@�ɶ�,����,a1,a2,a3,a4,a5,a6,b1,b2,b3,b4,b5,b6 FROM �Τ� WHERE id="&info(9),conn
if xx1=info(0) or xx2=info(0) then
diantiaook=1
else
diantiaook=0
end if
lxjwg1=rs("�Z�\")
lxjgj1=rs("����")
lxjfy1=rs("���s")
lxjyl1=rs("�Ȩ�")
egj1=rs("a1")+rs("a2")+rs("a3")+rs("a4")+rs("a5")+rs("a6")
efy1=rs("b1")+rs("b2")+rs("b3")+rs("b4")+rs("b5")+rs("b6")
if rs("�O�@")=true then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�Ш����m�\�O�@�A�ާ@�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("���A")="��" or rs("��O")<0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�A�w�g���F�r�����H�u���H�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if lxjyl1<=100000000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�A�S��100000000��Ȥl�A����o�X�o�ۯS�ޡA�٬O�h���I���a�I');location.href = 'javascript:history.go(-1)';</script>"
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
if rs("���H��")>=30 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�a�Ƨ@�ɡA���H�ƺ��A����ާ@�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if

if rs("����")<=5 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�����ۻݭn�԰�����5�ťH�W�~�i�H�ާ@�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if sj<100 then
	s=100-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('ĵ�i�G�е�"& s &"��A�o��,�i�O�ֵۡI');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
gw=left(to1,1)
bearcc=12345
if bearcc=12345 then
if IsNumeric(gw)=true then
select case info(8)
case "�Ԥh"
lxjgj1=lxjgj1*1
case "�ɲ��Ԥh"
lxjgj1=lxjgj1*2
case "���Z�Ԥh"
lxjgj1=lxjgj1*3
case "���ҾԤh"
lxjgj1=lxjgj1*4
case "���s�Ԥh"
lxjgj1=lxjgj1*5
case "�i�h"
lxjfy1=lxjfy1*1
case "�ʾԫi�h"
lxjfy1=lxjfy1*2
case "��«i�h"
lxjfy1=lxjfy1*3
case "���s�i�h"
lxjfy1=lxjfy1*3
case "����i�h"
lxjfy1=lxjfy1*4
case "�D�h"
lxjwg1=lxjwg1*1
case "���D�h"
lxjwg1=lxjwg1*1.5
case "���k�v"
lxjwg1=lxjwg1*2
case "���u�H"
lxjwg1=lxjwg1*2.5
case "���Ѯv"
lxjwg1=lxjwg1*3
end select
zz=split(to1,"��")
gw=zz(0)
if info(3)>gw*2 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('[�H�A�ثe�����ťΤ��ۥ���A�C2�����Ū��Ǫ��F]');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
  end if
rs.close
rs.open "select * from �Ǫ��� where �Ǫ�='"&to1&"'",conn
gwgj=rs("����")
gwfy=rs("���s")
gwtl=rs("��O")
killer=int((lxjgj1+lxjwg1+egj1+100000000)-gwfy*2)
if killer>=180000000 and Application("ljjh_bearcc")=1 then
randomize timer
kaa=int(rnd*99)+1
killer=180000000
killer=killer+kaa
end if
killergw=int(gwgj-(lxjwg1+efy1+lxjfy1))
if killer<=1 then
killer=int(rnd*99)+1
end if
if killergw<=1 then
killergw=int(rnd*99)+1
end if
gwclick="<font color=6FBFEF><a href=javascript:parent.sw('["+gw+"�ũǪ�]'); target=f2>"&aa&""& to1 & "</a></font>"
meclick="<a href=javascript:parent.sw('[" & info(0) & "]');parent.sws('[" & info(9) & "]'); target=f2>["&info(0)&"]</a>"
aa="<img src='../ico/"+gw+".gif'  border='0' width='36' height='36'>"
bb="<img src='../ico/"& info(10) &"-2.gif' width='36' height='36'>"
conn.execute "update �Ǫ��� set kill=kill-"  & killer & " where �Ǫ�='"&to1&"'"
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-100000000 where id="&info(9)
rs.close
rs.open "select * from �Ǫ��� where �Ǫ�='"&to1&"'",conn
if rs("kill")<0 then
jingye=int(rs("����"))
if rs("���H��")>10 then
sj=DateDiff("n",rs("�ާ@�ɶ�"),now())
if sj<5 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=5-sj
	Response.Write "<script language=JavaScript>{alert('���Ǫ��w�g���`�A�ЧA���W["& ss &"]����,�A�ާ@�I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
conn.execute "update �Ǫ��� set ���H��=0 where �Ǫ�='"&to1&"'"
end if
conn.execute "update �Ǫ��� set kill=��O,���H��=���H��+1,�ާ@�ɶ�=now() where �Ǫ�='"&to1&"'"
conn.execute "update �Τ� set allvalue=allvalue+"&jingye&" where �m�W='"&info(0)&"'"
stunt=info(0) & "<bgsound src=wav/jiuzhao3.wav loop=1><img src='img/18.gif'>�B�����O��"&aa&""&gwclick&"�ϥΤF����W����c�D�d�ǤU�Ӫ��m�@����<font color=6FBFEF>���믫�\</font>�A�������~�U�V�A�@���j���o�g�X�ӱ���"&gwclick&"�ͩR�ȡG" & killer & "�I,"&gwclick&"�ש󤣤�˦a�A�G���b"&meclick&"��W�I"&meclick&"��o"&jingye&"�I�g��ȡI"
call jiujian(stunt)
else
stunt=info(0) & "<bgsound src=wav/jiuzhao3.wav loop=1><img src='img/18.gif'>�B�����O��"&aa&""&gwclick&"�ϥΤF����W����c�D�d�ǤU�Ӫ��m�@����<font color=6FBFEF>���믫�\</font>�A�������~�U�V�A�@���j���o�g�X�ӱ���"&gwclick&"�ͩR�ȡG" & killer & "�I"&gwclick&"�İ_�����r��"&meclick&"�ͩR��:<font color=red>"&killergw&"</font>�I!"
conn.execute "update �Τ� set ��O=��O-"  & killergw & ",�ާ@�ɶ�=now() where �m�W='"&info(0)&"'"
call jiujian(stunt)
rs.close
rs.open "select ��O from �Τ� where �m�W='"&info(0)&"'",conn
if rs("��O")<0 then
Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="����"
sd(195)="�j�a"
sd(196)="6FBFEF"
sd(197)="6FBFEF"
sd(198)="��"
sd(199)="<font color=6FBFEF>�i�t�Ρj["&info(0)&"]</font>�j�L�Z�ǳy�ڥ����A�G���b"&gw&"�ũǪ���W!" 
sd(200)=0
Application("ljjh_sd")=sd
Application.UnLock

Response.Write "<script language=javascript>alert('�i"&info(0)&"�j�A��������A�G���b�Ǫ���W�F�I');</script>"
if info(4)=0 then 
call boot(info(0))
end if
end if
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('���ߡA�z�����믫�\�w�g�o�ۧ����I');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
rs.close
rs.open "select id,����,�|������,�Z�\,����,���H��,�O�@,���s,���A,��O,grade,����,b1,b2,b3,b4,b5,b6 FROM �Τ� WHERE �m�W='" & to1 & "'",conn
menpai=rs("����")
jhhy=rs("�|������")
lxjwg2=rs("�Z�\")
lxjgj2=rs("����")
lxjfy2=rs("���s")
idd=rs("id")
r=rs("����")
efy2=rs("b1")+rs("b2")+rs("b3")+rs("b4")+rs("b5")+rs("b6")
killnum=rs("���H��")
if xx1=to1 or xx2=to1 then
 if diantiaook=0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�A���i�H�����b��D�����H!');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
end if
if rs("�O�@")=true then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��西�b�����I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("���A")="��" or rs("��O")<-100 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�H���w�g���F�٥���');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("����")="�X�a" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('[" & to1 & "]�L�w�g�X�a�A����ާ@�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("����")<=10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('[" & to1 & "]���ŤӧC�A�g���_�A���ڡI');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
nei=100000000


killer=int((lxjwg1+lxjgj1+egj1+(denji1/r)*500000)-(lxjwg2+lxjfy2+efy2))
if killer<=100 then
randomize timer
killer=int(rnd*99)+1
end if
conn.execute "update �Τ� set ��O=��O-"  & killer & " where �m�W='" & to1 & "'"

conn.execute "update �Τ� set �Ȩ�=�Ȩ�-100000000 where id="&info(9)

e=""
rs.close
rs.open "select ��O FROM �Τ� WHERE �m�W='" & to1 & "'",conn
if rs("��O")<0 then
conn.execute "update �Τ� set ���H��=���H��+1,kill=kill+1,allvalue=allvalue+"&r&" where id="&info(9)
conn.execute "update �Τ� set ���O=100,�Z�\=100,����=100,���s=100,���A='��',�ާ@�ɶ�=now(),lastkick='"&info(0)&"',allvalue=allvalue-"&r&" where �m�W='" & to1 & "'"
'conn.execute "delete * from ���~ where �֦���="&ask&" and ����<>'�d��' and ����<>'�k��' and �˳�=false"
if xx1=info(0) or xx2=info(0) then
conn.execute "update ��D set �m�W1='�L',�m�W2='�L'"
end if
e1=","

rs.close
rs.open "select id,���~�W from ���~ where ( ����='�k��' or ����='����' or ����='���}' or ����='�Y��' or ����='�˹�' or ����='����' ) and �ƶq>0 and ��T��>0 and �˳�=1 and �֦���="&idd,conn
if not rs.eof or not rs.bof then
id=rs("id")
zb=rs("���~�W")
randomize timer
r=int(rnd*r)
conn.execute "update ���~ set ��T��=��T��-"&r&" where id="&id
rs.close
rs.open "select ��T�� from ���~ where id="&id,conn
if rs("��T��")<0 then 
call jiechu(idd,id)
end if
e1=",["&to1&"]���˳�["&zb&"]�l�a"&r&"�I!"
end if
if killnum>=10 then
rs.close
rs.open "select id,���~�W,���� from ���~ where ( ����='�k��' or ����='����' or ����='���}' or ����='�Y��' or ����='�˹�' or ����='����' ) and �ƶq>0  and �˳�=1 and �֦���="&idd,conn
if not rs.eof or not rs.bof then
id=rs("id")
	wupin=rs("���~�W")
	sm=rs("����")

Application.Lock
Application("ljjh_qiang")=1
Application.UnLock
conn.execute "insert into �H�R(����,�ɶ�,����,���]) values ('" & to1 & "',getdate(),'" & info(0) & "','���믫�\')"
e="�I�A" & to1 & "�C�C���ˤF�U�h  �q������W�S�֤F�@���j��,[" & info(0) & "]�԰��g��W�["&r&"�I," & to1 & "�l���԰��g��"&r&"�I!" & to1 & "�@���\�O�ɼo�A���b����....�����@�Ǫ��~�A�j�a�ַm!<a href='qiangb.asp?id=" & id&"' target='d'><img src='../hcjs/jhjs/001/"&sm&".gif' border='0' width='32' height='32'  align='top' alt='"&wupin&"'></a>" 
call jiechucolor(idd,id)
else
 e="�I"&e1&"" & to1 & "�C�C���ˤF�U�h  �q������W�S�֤F�@���j��,[" & info(0) & "]�԰��g��W�["&r&"�I," & to1 & "�l���԰��g��"&r&"�I!"
end if
else
 rs.close
rs.open "select id,���~�W,���� from ���~ where  �ƶq>0 and �˳�=false and ����<>'�d��' and �֦���="&idd,conn
if rs.eof or rs.bof then
           e="�I"&e1&"" & to1 & "�C�C���ˤF�U�h  �q������W�S�֤F�@���j��,[" & info(0) & "]�԰��g��W�["&r&"�I," & to1 & "�l���԰��g��"&r&"�I!" & to1 & "�@���\�O�ɼo�A���b����...."
else
id=rs("id")
	wupin=rs("���~�W")
	sm=rs("����")

Application.Lock
Application("ljjh_qiang")=1
Application.UnLock
conn.execute "insert into �H�R(����,�ɶ�,����,���]) values ('" & to1 & "',getdate(),'" & info(0) & "','���믫�\')"
e="�I�A" & to1 & "�C�C���ˤF�U�h  �q������W�S�֤F�@���j��,[" & info(0) & "]�԰��g��W�["&r&"�I," & to1 & "�l���԰��g��"&r&"�I!"&e1&"" & to1 & "�@���\�O�ɼo�A���b����....�����@�Ǫ��~�A�j�a�ַm!<a href='qiangb.asp?id=" & id&"' target='d'><img src='../hcjs/jhjs/001/"&sm&".gif' border='0' width='32' height='32'  align='top' alt='"&wupin&"'></a>" 
end if
end if
if info(4)=0 then 
call boot(to1)
end if

end if
stunt=info(0) & "�B�����O��<img src='img/18.gif'><font color=6FBFEF>" & to1 & "</font><bgsound src=wav/jiuzhao3.wav loop=1>�ϥΤF����W����c�D�d�ǤU�Ӫ��m�@����<font color=6FBFEF>���믫�\</font>�A�������~�U�V�A�@���j���o�g�X�ӱ���" & killer & e
call jiujian(stunt)
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('���ߡA�z�����믫�\�w�g�o�ۧ����I');location.href = 'javascript:history.go(-1)';</script>"
Response.End
else
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('���~�G�z���b�ϥΪ��F�C���򤣬O�����A�Ф����������A�F�C�`���Ghttp://xajh.us');location.href = 'javascript:history.go(-1)';</script>"
end if
sub jiujian(stunt)
Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="����"
sd(195)=info(0)
sd(196)="6FBFEF"
sd(197)="6FBFEF"
sd(198)="��"
sd(199)="<font color=6FBFEF>�i���믫�\�j"& stunt &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
end sub 

sub boot(to1)
Application.Lock
onlinelist=Application("ljjh_onlinelist"&session("nowinroom"))
dim newonlinelist()
useronlinename=""
onliners=0
js=1
for i=1 to UBound(onlinelist) step 6
if CStr(onlinelist(i+1))<>CStr(to1) then
onliners=onliners+1
useronlinename=useronlinename & " " & onlinelist(i+1)
Redim Preserve newonlinelist(js),newonlinelist(js+1),newonlinelist(js+2),newonlinelist(js+3),newonlinelist(js+4),newonlinelist(js+5)
newonlinelist(js)=onlinelist(i)
newonlinelist(js+1)=onlinelist(i+1)
newonlinelist(js+2)=onlinelist(i+2)
newonlinelist(js+3)=onlinelist(i+3)
newonlinelist(js+4)=onlinelist(i+4)
newonlinelist(js+5)=onlinelist(i+5)
js=js+6
else
kickip=onlinelist(i+2)
end if
next
useronlinename=useronlinename&" "
if kickip<>"" then
if onliners=0 then
dim listnull(0)
Application("ljjh_onlinelist"&session("nowinroom"))=listnull
else
Application("ljjh_onlinelist"&session("nowinroom"))=newonlinelist
end if
Application("ljjh_useronlinename"&session("nowinroom"))=useronlinename
Application("ljjh_chatrs"&session("nowinroom"))=onliners
else
Application.UnLock
Response.Redirect "manerr.asp?id=204&kickname=" & server.URLEncode(kickname)
end if
Application.UnLock
end sub
sub jiechu(idd,id)
rs.close
rs.open "SELECT ���� FROM ���~ WHERE �֦���=" & idd & " and �ƶq>0 and id=" & id,conn
leixing=rs("����")
select case leixing
case "�Y��"
conn.execute "update �Τ� set a1=0,b1=0 where id=" & idd
conn.execute "update ���~ set �˳�=false where id=" & id

case "����"

conn.execute "update �Τ� set a2=0,b2=0 where id=" & idd
conn.execute "update ���~ set �˳�=false where id=" & id

case "�k��"

conn.execute "update �Τ� set a3=0,b3=0 where id=" & idd
conn.execute "update ���~ set �˳�=false where id=" & id

case "���}"

conn.execute "update �Τ� set a4=0,b4=0 where id=" & idd
conn.execute "update ���~ set �˳�=false where id=" & id

case "�˹�"

conn.execute "update �Τ� set a5=0,b5=0 where id=" & idd
conn.execute "update ���~ set �˳�=false where id=" & id

case "����"

conn.execute "update �Τ� set a6=0,b6=0 where id=" & idd
conn.execute "update ���~ set �˳�=false where id=" & id

end select
conn.execute "delete  from  ���~  where id="& id
end sub
sub jiechucolor(idd,id)
rs.close
rs.open "SELECT ���� FROM ���~ WHERE �֦���=" & idd & " and �ƶq>0 and id=" & id,conn
leixing=rs("����")
select case leixing
case "�Y��"
conn.execute "update �Τ� set a1=0,b1=0 where id=" & idd
conn.execute "update ���~ set �˳�=false where id=" & id

case "����"

conn.execute "update �Τ� set a2=0,b2=0 where id=" & idd
conn.execute "update ���~ set �˳�=false where id=" & id

case "�k��"

conn.execute "update �Τ� set a3=0,b3=0 where id=" & idd
conn.execute "update ���~ set �˳�=false where id=" & id

case "���}"

conn.execute "update �Τ� set a4=0,b4=0 where id=" & idd
conn.execute "update ���~ set �˳�=false where id=" & id

case "�˹�"

conn.execute "update �Τ� set a5=0,b5=0 where id=" & idd
conn.execute "update ���~ set �˳�=false where id=" & id

case "����"

conn.execute "update �Τ� set a6=0,b6=0 where id=" & idd
conn.execute "update ���~ set �˳�=false where id=" & id

end select
end sub
%>
<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Session("ljjh_inthechat")<>"1" then 
	Response.Write "<script Language=Javascript>alert('�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
at1=request.form("at1")
at2=request.form("at2")
at3=request.form("at3")
to1=request.form("to1")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(to1)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�ҧ������H���b��ѫǡI');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if Application("jingda")=1 then
	Response.Write "<script Language=Javascript>alert('�ثe�ѩ�H�֪���]�t�γQ�]�m���T�����A�A�n�Q���[�Ч�x���H���}�ҡI');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
from1=info(0)
stunt=""
stunt1=""
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select grade,���A,�ާ@�ɶ�,���H��,��O,����,�O�@ from �Τ� where id="&info(9),conn
if rs("���A")="��" or rs("��O")<0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�A�w�g���F�r�����H�u���H�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('ĵ�i�G�е�"& s &"��A�o��,�i�O�ֵۡI');location.href = 'javascript:history.go(-1)';</script>"
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
	Response.Write "<script Language=Javascript>alert('�s��޻ݭn�԰�����5�ťH�W�~�i�H�ާ@�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("�O�@")=true then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�Ш����m�\�O�@�A�ާ@�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
rs.open "select �|������,����,��O,���A,�O�@,���� from �Τ� where �m�W='" & to1 & "'",conn
jhhy=rs("�|������")
mp=info(5)
if rs("���A")="��" or rs("��O")<0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�A�w�g���F�r�����H�u���H�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("�O�@")=true then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('["&to1&"]�A�b�m�\�O�@���A����ާ@�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("����")<=10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('["&to1&"]����֦���s��A���Τ@�o�򭫪��ۦ��a�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
rs.open "SELECT �Z�\,���O FROM �Z�\ WHERE �Z�\='" & at1 & "' or �Z�\='" & at2 & "' or �Z�\='" & at3 & "' and ����='" & mp & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�A�èS���o�˪��Z�\�r�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("�Z�\")=at1 then
nei1=abs(rs("���O"))
end if
if rs("�Z�\")=at2 then
nei2=abs(rs("���O"))
end if
if rs("�Z�\")=at3 then
nei3=abs(rs("���O"))
end if
nei=nei1+nei2+nei3
rs.close
rs.open "select �Z�\,����,���O from �Τ� where id="&info(9),conn
lxjwg1=rs("�Z�\")
lxjgj1=rs("����")
if rs("���O")<nei  then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���O�����A�o���F�s��ޡI');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
rs.open "select �Z�\,����,��O from �Τ� where �m�W='" & to1 & "'",conn
lxjwg2=rs("�Z�\")
lxjgj2=rs("����")
killer=int(((lxjwg1+lxjgj1)-(lxjwg2+lxjgj2)+nei/10)/5)
'���ˤp��100
if killer<=100 then
	randomize timer
	killer=int(rnd*99)+1
end if
conn.execute "update �Τ� set ���O=���O-" & nei & ",�ާ@�ɶ�=now() where id="&info(9)
conn.execute "update �Τ� set ��O=��O-"  & killer & " where �m�W='" & to1 & "'"
e=""
if rs("��O")<-100 then
	conn.execute "update �Τ� set ���H��=���H��+1 where id="&info(9)
	conn.execute "update �Τ� set ���A='��' where �m�W='" & to1 & "'"
	conn.execute "insert into �H�R(����,�ɶ�,����,���]) values ('" & to1 & "',now(),'" & info(0) & "','�s���')"
	e="�I�A" & to1 & "�C�C���ˤF�U�h  �q������W�S�֤F�@���j��"
	call boot(to1)
end if
stunt=info(0) & "�B�����O��" & to1 & "�ϥΤF�i" & at1 & "+" & at2 & "+" & at3 & "�j���@�s����R�s��ޡA����" & killer & e

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
sd(196)="FFD060"
sd(197)="FFD060"
sd(198)="��"
sd(199)="<font color=red>�i�s��ޡj"& stunt &"</font>"& t
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script Language=Javascript>alert('���ߡA�z���s��ޤw�g�o�ۧ����I');location.href = 'javascript:history.go(-1)';</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
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
%>
 
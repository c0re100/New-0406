<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if Session("ljjh_inthechat")<>"1" then 
	Response.Write "<script Language=Javascript>alert('�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Response.Expires=0
useronlinename=Application("ljjh_useronlinename"&session("nowinroom"))
at=request.form("at")
to1=request.form("to1")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(to1)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�ҧ������H���b��ѫǡI');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
from1=info(0)
compact=""
compact1=""
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �ާ@�ɶ�,���A,���H��,����,�O�@,�G�B from �Τ� where id="&info(9),conn
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
if rs("���A")="��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�A�w�g���F�r�����H�u���H�I');location.href = 'javascript:history.go(-1)';</script>"
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
	Response.Write "<script Language=Javascript>alert('�ҩd�X��޻ݭn�԰�����5�ťH�W�~�i�H�ާ@�I');location.href = 'javascript:history.go(-1)';</script>"
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
peiou=rs("�G�B")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(PEIOU)&" ")=0 Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�A���G�B["&peiou&"]�S���b��ѫǤ�����ާ@�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
rs.open "select ����,���A,�O�@,�|������,���� from �Τ� where �m�W='" & to1 & "'",conn
if rs("���A")="��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���w�g���F�A�u���H�����H�I');location.href = 'javascript:history.go(-1)';</script>"
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
jhhy=rs("�|������")
ntnt=rs("����")
mp=info(5)
if rs("����")<=2 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('["&to1&"]����֦���s��A���Τ@�o�򭫪��ۦ��a�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
rs.open "SELECT * FROM �X��� WHERE �m�W�k='" & info(0) & "' or �m�W�k='" & info(0) & "' and �X���='" & at & " '",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�A�èS��["&at&"]�o�˪��ҩd�X��ާr�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
nei=abs(rs("���O"))
rs.close
rs.open "select �Z�\,���� from �Τ� where �m�W='" & peiou & "'",conn
	htwg1=rs("�Z�\")
	htgj1=rs("����")
rs.close	
rs.open "select ���O,�Z�\,���� from �Τ� where id="&info(9),conn
if rs("���O")<nei then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���O�����A�o���F�o�ۦX��ޡI');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if

htwg2=rs("�Z�\")
htgj2=rs("����")
rs.close
rs.open "select �Z�\,���s,��O from �Τ� where �m�W='" & to1 & "'",conn
towg=rs("�Z�\")
tofy=rs("���s")
killer=int((((htwg1+htwg2+htgj1+htgj2)-towg-tofy)+nei/10)/7)
'�p�G���ˤO�p��100�H����
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
	conn.execute "insert into �H�R(����,�ɶ�,����,���]) values ('" & to1 & "',now(),'" & from1 & "','�X���')"
	call boot(to1)
	e="�I�A" & to1 & "�C�C���ˤF�U�h  �q������W�S�֤F�@���j��"
end if
compact="["&from1 & "]�B��" & nei & "�I���O�P{" & peiou & "}�@�_��(" & to1 & ")�ϥΤF�W��" & at & "���ҩd�X��ޡA����" & killer & e
rs.close
set rs=nothing
conn.close
set conn=nothing
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
sd(196)="B0D0E0"
sd(197)="B0D0E0"
sd(198)="��"
sd(199)="<font color=B0D0E0>�i�X��ޡj"& compact &"</font>"& t
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script Language=Javascript>alert('���ߡA�z���ҩd�X��ޤw�g�����I');location.href = 'javascript:history.go(-1)';</script>"
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
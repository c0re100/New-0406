<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Session("ljjh_inthechat")<>"1" then
Response.Write "<script Language=Javascript>alert('�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
at=request.form("at")
to1=request.form("to1")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(to1)&" ")=0 then
Response.Write "<script Language=Javascript>alert('�ҧ������H���b��ѫǡI');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
from1=info(0)
pet=""
pet1=""
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select grade,���A,�ާ@�ɶ�,���H��,����,�O�@ from �Τ� where id="&info(9),conn

if rs("���A")="��" then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�A�w�g���F�r�����H�u���H�I');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if sj<10 then
s=60-sj
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
if rs("����")<=3 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�d�������ݭn�԰�����3�ťH�W�~�i�H�ާ@�I');location.href = 'javascript:history.go(-1)';</script>"
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
rs.open "select �|������,����,���s,����,��O,��O�[,����,�O�@ from �Τ� where �m�W='" & to1 & "'",conn
jhhy=rs("�|������")
mp=info(5)
togj=rs("����")
tofy=rs("���s")
dengji=rs("����")
totl=rs("��O")
tlj=(dengji*1500+5260)+rs("��O�[")
if rs("�O�@")=true then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('���b�m�\�O�@���A����ާ@�I');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
if rs("����")<=2 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('��謰����s��A�٬O���Τ@�o�򭫪��ۦ��a�I');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
rs.close
rs.open "SELECT * FROM �d�����A WHERE user='" & from1 & "'",conn
if rs.eof or rs.bof then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�A�èS���d���A�������r�I');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
adde=rs("����id")
att=abs(rs("����"))
lx=rs("�d������")
say=rs("����")
jn=rs("�ޯ�")
petname=rs("�W�r")
ap=abs(rs("�g��"))
zou=abs(rs("����"))
fangyu=abs(rs("���s"))
Select Case at
case "���q����"
kills=att*(ap+zou)
randomize timer
xqz=int(rnd*100)+10
xx=int(rnd*10)
kills=int(kills-tofy*togj)
if jn="��O��" then
conn.execute "update �Τ� set ��O=��O+"&xqz&" where id="&info(9)
conn.execute "update �Τ� set ��O=��O-"&xqz&" where �m�W='"&to1&"'"
aa="�l��<font color=9FDF70>"&to1&"</font><font color=red>"&xqz&"</font>�I��O"
end if
if jn="���O��" then
conn.execute "update �Τ� set ���O=���O+"&xqz&" where id="&info(9)
conn.execute "update �Τ� set ���O=���O-"&xqz&" where �m�W='"&to1&"'"
aa="�l��<font color=9FDF70>"&to1&"</font><font color=red>"&xqz&"</font>�I���O"
end if
if jn="���s��" then
conn.execute "update �Τ� set ���s=���s+"&xqz&" where id="&info(9)
conn.execute "update �Τ� set ���s=���s-"&xqz&" where �m�W='"&to1&"'"
aa="����<font color=9FDF70>"&to1&"</font><font color=red>"&xqz&"</font>�I���s�O"
end if
if jn="�Z�\��" then
conn.execute "update �Τ� set �Z�\=�Z�\+"&xqz&" where id="&info(9)
conn.execute "update �Τ� set �Z�\=�Z�\-"&xqz&" where �m�W='"&to1&"'"
aa="�Ƿ|<font color=9FDF70>"&to1&"</font><font color=red>"&xqz&"</font>�I�Z�\"
end if
if jn="�k�O��" then
conn.execute "update �Τ� set �k�O=�k�O+"&xqz&" where id="&info(9)
conn.execute "update �Τ� set �k�O=�k�O-"&xqz&" where �m�W='"&to1&"'"
aa="�l��<font color=9FDF70>"&to1&"</font><font color=red>"&xqz&"</font>�I�k�O"
end if


if kills>10 and kills<tlj then
conn.execute "update �d�����A set ��O=��O-" & int(togj/fangyu) & ",�g��=�g��+10 where user='" & from1 & "'"
conn.execute "update �Τ� set �ާ@�ɶ�=now() where id="&info(9)
conn.execute "update �Τ� set ��O=��O-"  & kills & " where �m�W='" & to1 & "'"
rs.close
rs.open "select * from �Τ� where �m�W='" & to1 & "'",conn
if rs("��O")<-100 then
conn.execute "update �Τ� set ���H��=���H��+1 where id="&info(9)
conn.execute "update �Τ� set ���A='��' where �m�W='" & to1 & "'"
conn.execute "insert into �H�R(����,�ɶ�,����,���]) values ('" & to1 & "',now(),'" & info(0) & "','�d������')"
e="�I�A" & to1 & "�C�C���ˤF�U�h  �q������W�S�֤F�@���j��"
call boot(to1)
end if
pet="<a href=javascript:parent.sw('[" & from1 & "]'); target=f2><font color=red>"&from1&"</font></a>�l��X<img src='../myhome/sheep/image/"&say&".gif' width=49 height=46>[" & lx & "]�ϥX�����q����������" & to1 & "�A�d���Q�����A��O���C" & int(togj/fangyu) & "�I�A�r��" & to1 & "��O" & kills & e
elseif kills>=tlj then
conn.execute "update �d�����A set ��O=��O-" & int(togj/fangyu) & ",�g��=�g��+10 where user='" & from1 & "'"
conn.execute "update �Τ� set �ާ@�ɶ�=now() where id="&info(9)
conn.execute "update �Τ� set ��O=��O/2,����=����-1000,���s=���s-1000 where �m�W='" & to1 & "'"
rs.close
rs.open "SELECT ��O FROM �d�����A WHERE user='" & from1 & "'",conn
if rs("��O")>0 then
pet="<a href=javascript:parent.sw('[" & from1 & "]'); target=f2><font color=red>"&from1&"</font></a>�l��X<img src='../myhome/sheep/image/"&say&".gif' width=49 height=46>[" & lx & "]�ϥX--[���q����]--����<font color=9FDF70>" & to1 & "</font>�A�d���Q�����A��O���C<font color=red>" & int(togj/fangyu) & "</font>�I�A" & to1 & "�����B���s�U���1000�I�A��O���" & totl/2& "�I"
else
pet="<a href=javascript:parent.sw('[" & from1 & "]'); target=f2><font color=red>"&from1&"</font></a>���d��<img src='../myhome/sheep/image/"&say&".gif' width=49 height=46>[" & lx & "]�g�����_<font color=9FDF70>" & to1 & "</font>���@�y�r���A�ש�C�C�C�C���ˤF�U�h  �q������S�֤F�@���d���I"
conn.execute "Delete * From �d�����A Where user='" & from1 & "'"
'conn.execute "Delete * From �ޯ�C�� Where �֦���='" & from1 & "'"
conn.execute "Delete * From �D��C�� Where �֦���='" & from1 & "'"
end if
else
if xx=9 or xx=7 then
pet="<a href=javascript:parent.sw('[" & from1 & "]'); target=f2><font color=red>"&from1&"</font></a>�l��X<img src='../myhome/sheep/image/"&say&".gif' width=49 height=46>[" & lx & "]�ϥX--[���q����]--����<font color=red>" & to1 & "</font>�C<font color=red>" & to1 & "</font>�]�ϥX�����ѼơA�԰��i�檺���`�E�P�A�̫����襴������A���L�l���I"
elseif xx=8 then
sql="update �d�����A set ��O=��O-"& xqz &",���s=���s-"& xqz*2 &" where user='" & from1 & "'"
conn.execute sql
rs.close
rs.open "SELECT ��O FROM �d�����A WHERE user='" & from1 & "'",conn
if rs("��O")<0 then
pet="<a href=javascript:parent.sw('[" & from1 & "]'); target=f2><font color=red>"&from1&"</font></a>���d��<img src='../myhome/sheep/image/"&say&".gif'width=49 height=46>[" & lx & "]�g�����_<font color=9FDF70>" & to1 & "</font>���@�y�r���A�ש�C�C�C�C���ˤF�U�h  �q������S�֤F�@���d���I"
conn.execute "Delete * From �d�����A Where user='" & from1 & "'"
'conn.execute "Delete * From �ޯ�C�� Where �֦���='" & from1 & "'"
conn.execute "Delete * From �D��C�� Where �֦���='" & from1 & "'"
else
pet="<a href=javascript:parent.sw('[" & from1 & "]'); target=f2><font color=red>"&from1&"</font></a>�l��X<img src='../myhome/sheep/image/"&say&".gif' width=49 height=46>[" & lx & "]�ϥX--[���q����]--����<font color=red>" & to1 & "</font>�C<font color=red>" & to1 & "</font>�j�ۡG�֮a������d���]�������ڡH�����ڪ����h�������������ۡA" & petname & "���ۧ��ڰk�]�F�A�l����O<font color=red>" & xqz & "</font>�I�A���s�l��<font color=red>" & 2*xqz & "</font>�I�I"
end if
else
conn.execute "update �Τ� set �ާ@�ɶ�=now() where id="&info(9)
conn.execute "update �Τ� set ��O=��O-"  & xqz & " where �m�W='" & to1 & "'"
rs.close
rs.open "select * from �Τ� where �m�W='" & to1 & "'",conn
if rs("��O")<-100 then
conn.execute "update �Τ� set ���H��=���H��+1 where id="&info(9)
conn.execute "update �Τ� set ���A='��' where �m�W='" & to1 & "'"
conn.execute "insert into �H�R(����,�ɶ�,����,���]) values ('" & to1 & "',now(),'" & info(0) & "','�d������')"
e="�I�A" & to1 & "�C�C���ˤF�U�h  �q������W�S�֤F�@���j��"
call boot(to1)
end if
pet="<a href=javascript:parent.sw('[" & from1 & "]'); target=f2><font color=red>"&from1&"</font></a>�l��X<img src='../myhome/sheep/image/"&say&".gif' width=49 height=46>[" & lx & "]�ϥX--[���q����]--����<font color=9FDF70>" & to1 & "</font>�A"&aa
end if
end if
case "������"
randomize timer
xqz=int(rnd*50)+10
rs.close
sql="SELECT * FROM �d�����A WHERE user='" & to1 & "'"
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
pet="<a href=javascript:parent.sw('[" & from1 & "]'); target=f2><font color=red>"&from1&"</font></a>�ϥX�����������ѡA�]�����S���d���C"
else
id=rs("����id")
tosay=rs("����")
petname1=rs("�W�r")
kills=int(rs("���s")+rs("����")-fangyu-att-(ap+zou)/10)
if kills<10 then
kills=xqz
end if
conn.execute "update �d�����A set �g��=�g��+10 where user='" & from1 & "'"
conn.execute "update �d�����A set ��O=��O-"  & kills & " where user='" & to1 & "'"
e=""
if rs("��O")-kills<=0 then
conn.execute "update �d�����A set �g��=�g��+10 where user='" & from1 & "'"

conn.execute "update �d�����A set ��O=��O-"  & kills & " where user='" & to1 & "'"

e="�A<font color=red>" & to1 & "</font>���d��<img src='../myhome/sheep/image/"&tosay&".gif' width=49 height=46>[" & petname1 & "]�C�C���ˤF�U�h  �q������S�֤F�@���d���I"

end if
pet="<a href=javascript:parent.sw('[" & from1 & "]'); target=f2><font color=red>"&from1&"</font></a>�l��X<img src='../myhome/sheep/image/"&say&".gif' width=49 height=46>[" & lx & "]�ϥX--[������]--����<font color=red>" & to1 & "</font>���d���A�r��<img src='../myhome/sheep/image/"&tosay&".gif' width=49 height=46>[" & petname1 & "]��O<font color=red>" & kills &"</font>"& e
if rs("��O")<=0 then
conn.execute "Delete * From �d�����A Where user='" & to1 & "'"
conn.execute "Delete * From �D��C�� Where �֦���='" & to1 & "'"
end if
end if
End Select
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
sd(194)=to1
sd(195)=info(0)
sd(196)="9FDF70"
sd(197)="9FDF70"
sd(198)="��"
sd(199)="<font color=9FDF70>�i�d�������j"& pet &"</font>"& t
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('���ߡA�z���d���w�g�o�ۧ����I');location.href = 'javascript:history.go(-1)';</script>"

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

<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('���ܡG�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');window.close();</script>"
	response.end
end if
if Application("ljjh_kl")=0 then
	Response.Write "<script Language=Javascript>alert('���ܡG�A�p�l�O���O�S�Ƨ�Ƨr�A���̦����ǡH�I');</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
conn.execute "update �Τ� set ��O=��O-"&Application("ljjh_kl")&" where id="&info(9)
rs.open "select ��O FROM �Τ� WHERE id="&info(9),conn
if rs("��O")<-100 then
	conn.execute "update �Τ� set ���A='��',lastkick='����' where id="&info(9)
        
	call boot(info(0))
	conn.execute "insert into �H�R(����,�ɶ�,����,���]) values ('" & info(0) & "',now(),'����','�]�i�����ǩ��묹')"
	kl="<img src='pic/kl.gif'>�Ӥ����F�A["&info(0)&"]���F�O�@�j�a���w���A�٨����q�A�Q���ǫr���C�s�K���A��¼¼�F   "
else
	conn.execute "update �Τ� set �Ȩ�=�Ȩ�+"& Application("ljjh_kl")*55 &" where �m�W='" & info(0) & "'"
	kl="��call,�^���A�u�O�^���A["&info(0)&"]�P���Ǥj���X��  ���ǲש�ˤU�F�A�ӥL�ۤv�]�˪������A�x�����y<img src='img/251.gif'>"&Application("ljjh_kl")*55&"��I"
	Application.Lock
	Application("ljjh_kl")=0
	Application.UnLock
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application.Lock
Application("ljjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="����"
sd(195)="�j�a"
sd(196)="9FDF70"
sd(197)="9FDF70"
sd(198)="��"
sd(199)="<font color=9FDF70>�i��������j</font>"&kl
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
'��H
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
if onliners=0 then
dim listnull(0)
Application("ljjh_onlinelist"&session("nowinroom"))=listnull
else
Application("ljjh_onlinelist"&session("nowinroom"))=newonlinelist
end if
Application("ljjh_useronlinename"&session("nowinroom"))=useronlinename
Application("ljjh_chatrs"&session("nowinroom"))=onliners
Application.UnLock
end sub
%>
 
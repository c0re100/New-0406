<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
name=request("name")
my=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="select ���A from �Τ� where �m�W='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
response.write "�A���O���򤤤H�A����ϤH"
conn.close
response.end
end if 
if rs("���A")<>"��" and rs("���A")<>"�c" then
randomize timer
r=int(rnd*8)
if r=1 then
sql="update �Τ� set ���A='���`' where �m�W='" & name & "'"
conn.execute sql
msg="<font color=FFD7F4>�i�T���j</font><font color=FFD7F4>"&info(0)&"</font>���ۤj�M���J�c�̧�<font color=FFD7F4>" & name & "</font>���ϤF�X�ӡI"
else
sql="update �Τ� set ���A='��' where �m�W='" & my & "'"
conn.execute sql
msg="<font color=FFD7F4>�i�T���j</font><font color=FFD7F4>"&info(0)&"</font>���ۤj�M���J�c�̷Q��<font color=FFD7F4>" & name & "</font>���G�ۤv�Q��F�I"
call boot(info(0))
end if
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
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="��"
sd(199)="<font color=FFD7F4>"& msg &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
set rs=nothing
conn.close
set conn=nothing
end if
Response.Redirect "listlao.asp"
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
Response.Redirect "chaterr.asp?id=001"
Application.UnLock
end sub
%>
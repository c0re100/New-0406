<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%
if not IsArray(Session("info")) then Response.Redirect "error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
nickname=info(0)
conn.Execute "update 用戶 set 寶物='無' where  姓名='" & nickname & "'"
if Session("ljjh_inthechat")="1" then
nickname=info(0)
onlinelist=Application("ljjh_onlinelist"&session("nowinroom"))
dim newonlinelist()
useronlinename=""
onliners=0
js=1
ubl=UBound(onlinelist)
for i=1 to ubl step 6
if CStr(onlinelist(i+1))<>CStr(nickname) then
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
savetime=onlinelist(i+5)
end if
next
useronlinename=useronlinename&" "
if savetime<>"" then
if onliners=0 then
dim listnull(0)
Application.Lock
Application("ljjh_onlinelist"&session("nowinroom"))=listnull
else
Application("ljjh_onlinelist"&session("nowinroom"))=newonlinelist
end if
Application("ljjh_useronlinename"&session("nowinroom"))=useronlinename
Application("ljjh_chatrs"&session("nowinroom"))=onliners
Application.UnLock
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
t=s & ":" & f & ":" & m
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
sd(194)=nickname
sd(195)="mingdan"
sd(196)="DFF0C0"
sd(197)="DFF0C0"
sd(198)="對"
sd(199)="<font color=F0704F>【公告】" & nickname & "</font><font color=DFF0C0>大俠離開江湖空間,回到登錄頁面準備重進  </font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
end if
end if
Session("ljjh_inthechat")="0"
info(0)=""
Application.UnLock
Session.Abandon
Session("info")=""
Session("ljjh_name")=""
Response.Redirect "index.asp"
response.end
%>
 
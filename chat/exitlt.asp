<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if Session("ljjh_inthechat")<>"1" then Response.Redirect "close.asp"
Session("ljjh_inthechat")="0"
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
t=s & ":" & f & ":" & m
sj=n & "-" & y & "-" & r & " " & t
ip=Request.ServerVariables("REMOTE_ADDR")
Application.Lock
onlinelist=Application("ljjh_onlinelist"&session("nowinroom"))
dim newonlinelist()
useronlinename=""
onliners=0
js=1
ubl=UBound(onlinelist)
for i=1 to ubl step 6
if CStr(onlinelist(i+1))<>CStr(info(0)) then
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
Application("ljjh_onlinelist"&session("nowinroom"))=listnull
else
Application("ljjh_onlinelist"&session("nowinroom"))=newonlinelist
end if
Application("ljjh_useronlinename"&session("nowinroom"))=useronlinename
Application("ljjh_chatrs"&session("nowinroom"))=onliners
if Instr(1,Application("ljjh_ying"),"|"& info(0) & "|")=0  then
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
j=1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)=info(0)
sd(195)="mingdan"
sd(196)="DFF0C0"
sd(197)="DFF0C0"
sd(198)="對"
sd(199)="<font color=F06F4F>【公告】</font>“各位大俠，" & info(0) & "出去買點東西，一會兒見！"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
end if
End if
Application.UnLock
session("nowinroom")=0
if savetime<>"" then
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("ljjh_usermdb")
addvalue=DateDiff("n",savetime,sj)
conn.execute  "update 用戶 set allvalue=allvalue+"&addvalue&",mvalue=mvalue+"&addvalue&",lasttime='"&sj&"',lastip='"&ip&"' where id="&info(9)
conn.close
set conn=nothing
end if
Response.Redirect "close.asp"%>
 
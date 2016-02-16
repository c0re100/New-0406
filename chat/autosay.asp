<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.Expires=0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
ljjh_name=info(0)
Response.Write "<script language=javascript>if(window==window.top){top.location.href='chaterr.asp';}</script>"
if info(0)="" then Response.Redirect "chaterr.asp"
useronlinename=Application("ljjh_useronlinename"&session("nowinroom"))
if info(0)="" or Session("ljjh_inthechat")<>"1" or Instr(LCase(useronlinename),LCase(" "&info(0)&" "))=0 then
Session("ljjh_inthechat")="0"
Response.Redirect "chaterr.asp?id=001"
end if
if DateDiff("n",Session("ljjh_savetime"),now())>=15 then
	Response.Write "<script Language=Javascript>parent.clsok=1;if(parent.chatcls>1){parent.qp();}parent.f3.location.href='savevalue.asp';</script>"
end if
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
sj=n & "-" & y & "-" & r & " " & s & ":" & f & ":" & m
t="..."
Session("ljjh_lasttime")=sj
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
sd(194)="消息"
sd(195)=info(0)
sd(196)="FFCFCF"
sd(197)="FFCFCF"
sd(198)="對"
sd(199)="<font color=FFCFCF>【自動泡點】["&info(0)&"]說：人不在，泡點中  </font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.Lock
%>
<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(0)="" or Session("ljjh_inthechat")<>"1" then Response.Redirect "chaterr.asp?id=001"
song=Request.Form("song")
loopok=Request.Form("loopok")
linkurl="mid/" & song &".mid"
Response.write "<html><head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5"> </head><body oncontextmenu=self.event.returnValue=false>"
if Request.Form("play")="播放" then
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
Session("ljjh_lasttime")=sj
Response.Write "<bgsound src=" & chr(34) & linkurl & chr(34) & " loop=" & chr(34) & loopok & chr(34) & "><script Language=JavaScript>parent.f2.startnosay();</script>"
end if
Response.Write "</body></html>"%> 

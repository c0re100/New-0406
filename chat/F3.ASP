<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
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
online=Application("ljjh_onlinelist"&session("nowinroom"))
%>
<script Language="JavaScript">
if(parent.document.URL.indexOf("file:")>=0||parent.f2.document.URL.indexOf("file:")>=0){parent.location.href='chaterr.asp?id=001';}
parent.md1();
parent.f3.document.writeln("<div align=\"center\"><font color=\"#FFFF66\"<%if len(Application("ljjh_chatroomname"&session("nowinroom")))>6 then Response.Write " style='font-size:11pt'"%>><%=Application("ljjh_chatroomname"&session("nowinroom"))%></font><hr size=1 color=FFFF00>");
parent.f3.document.writeln("<font color=yellow style=\"font-size:11pt\"><%=Application("ljjh_chatrs"&session("nowinroom"))%></font><font color=white style=\"font-size:11pt\"><span class=p9></span>¤H¦b½u<br></div>");
parent.md2();
<%

for i=1 to UBound(online)   step 6
if Instr(1,Application("ljjh_ying"),"|"& online(i+1) & "|")=0 then
Response.Write "parent.tma("&chr(34)& online(i+3) &chr(34)&");"& chr(13) & chr(10)
end if
next %>
parent.mdsm();
parent.md3();
</script>
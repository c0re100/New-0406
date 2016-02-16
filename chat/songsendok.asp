<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<!--#INCLUDE FILE="midsound.asp" -->
<%Response.Expires=0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
useronlinename=Application("ljjh_useronlinename"&session("nowinroom"))
if info(0)="" or Session("ljjh_inthechat")<>"1" or Instr(useronlinename," "&info(0)&" ")=0 then Response.Redirect "chaterr.asp?id=001"
towho=Trim(Request.Form("towho"))
if Instr(useronlinename," "&towho&" ")=0 then Response.Redirect "../error.asp?id=410&name=" & server.URLEncode(towho)
song=Request.Form("song")
infofo=Left(Trim(Request.Form("infofo")),100)
infofo=Server.HTMLEncode(Replace(infofo,"%%",mid(song)))
linkurl="mid/"& song &".mid"
linkurl2="mid/" & Server.URLEncode(song) &".mid"
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
wbq=Application("ljjh_webicq")
wbqub=UBound(wbq)
if wbqub>0 then
Dim wbqnew()
j=1
for i=1 to wbqub step 4
	if DateDiff("n",wbq(i),sj)<=10 then
		Redim Preserve wbqnew(j),wbqnew(j+1),wbqnew(j+2),wbqnew(j+3)
		wbqnew(j)=wbq(i)
		wbqnew(j+1)=wbq(i+1)
		wbqnew(j+2)=wbq(i+2)
		wbqnew(j+3)=wbq(i+3)
		j=j+4
	end if
next
	if j>=4 then
		wbq=wbqnew
	else
		Dim wbqnull(0)
		wbq=wbqnull
	end if
	wbqub=UBound(wbq)
end if
Redim Preserve wbq(wbqub+1),wbq(wbqub+2),wbq(wbqub+3),wbq(wbqub+4)
wbq(wbqub+1)=sj
wbq(wbqub+2)=towho
wbq(wbqub+3)=info(0)
wbq(wbqub+4)="<div align=center><bgsound src="& linkurl &"></div>" & infofo
wbqub=wbqub+4
webicqname=""
for i=1 to wbqub step 4
	webicqname=webicqname & " " & wbq(i+1)
next
webicqname=webicqname&" "
Application.Lock
Application("ljjh_webicq")=wbq
Application("ljjh_webicqname")=webicqname
Application.UnLock
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
sd(195)="大家"
sd(196)="FFD060"
sd(197)="FFD060"
sd(198)="對"
sd(199)="<font color=FFD060>【靈劍點歌】</font><font color=FFD060>"&info(0)&"</font>偷偷給<font color=red>["&towho&"]</font>點了一首<font color=FFD060>["&mid(song)&"]</font>並說道： <font color=FFD060>"&infofo&"</font>" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
if chatbgcolor="" then chatbgcolor="008888"%>
<html>
<head>
<title>好歌贈好友</title>
<meta http-equiv='content-type' content='text/html; charset=big5'>
<link rel="stylesheet" href="readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#660000" leftmargin="0" bgproperties="fixed" text="#FFFFFF" topmargin="0">
<div align="center"><font color="#FFFF00" style="font-size:12pt">發送完畢</font></div>
<hr size=1 color=FFFF00><br>
<p>已經將：</p>
<p><font color=FFFF00><%=mid(song)%></font></p>
<p>發送給：</p>
<p><font color=FFFF00><%=towho%></font></p>
<p>祝福話語：</p>
<p><font color=FFFF00><%=infofo%></font></p>
<input type="button" name="abort" value="返回" onClick="javascript:history.go(-2)">
</body>
</html> 
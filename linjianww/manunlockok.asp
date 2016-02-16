<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
useronlinename=Application("ljjh_useronlinename")
nickname=info(0)
inthechat=Session("ljjh_inthechat")
userip=Request.ServerVariables("REMOTE_ADDR")
if info(2)<10   then Response.Redirect "../error.asp?id=900"
if inthechat<>"1"  then Response.Redirect "manerr.asp?id=211"
unlockip=Server.HTMLEncode(Trim(Request.Form("unlockip")))
if unlockip="" then Response.Redirect "manerr.asp?id=218"
unlockwhy=Server.HTMLEncode(Trim(Request.Form("unlockwhy")))
logok=Trim(Request.Form("logok"))
if logok<>"0" then logok="1"
if unlockwhy="" then Response.Redirect "manerr.asp?id=216"
if len(unlockwhy)>60 then unlockwhy=Left(unlockwhy,60)
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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="SELECT ip FROM iplocktemp WHERE ip='" & unlockip & "'"
rs.open sql,conn,1,1
if rs.Eof and rs.Bof then
rs.close
conn.close
set rs=nothing
set conn=nothing
Response.Redirect "manerr.asp?id=217"
end if
rs.close
set rs=nothing
sql="DELETE FROM iplocktemp WHERE ip='" & unlockip & "'"
conn.Execute sql
Function SqlStr(data)
SqlStr="'" & Replace(data,"'","''") & "'"
End Function
unlocklog="解鎖IP：<font color=red>" & unlockip & "</font>！<br><font color=FFD7F4>【原因：" & unlockwhy & "】</font>"
if logok="1" then
sql="INSERT INTO logdata (logtime,name,ip,opertion) VALUES ("
sql=sql & SqlStr(sj) & ","
sql=sql & SqlStr(nickname) & ","
sql=sql & SqlStr(userip) & ","
sql=sql & SqlStr(unlocklog) & ")"
conn.Execute sql
end if
conn.close
set conn=nothing
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
sd(194)=nickname
sd(195)="大家"
sd(196)="FFD7F4"&cstr(chatroomsn)
sd(197)="FFD7F4"
sd(198)="對"
sd(199)="<font color=FFD7F4>【解鎖】</font><font color=FFD7F4><font color=FFD7F4>" & nickname & "</font>解鎖IP：<font color=red>" & unlockip & "</font>！ 原因：" & unlockwhy & " </font><font class=t></font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock%><html>
<head>
<title>解鎖ＩＰ</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: 新細明體}
.c {  font-family: 新細明體; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#FFFFFF" class=p150 background="0064.jpg">
<div align="center">
  <h1><font color="#FF0000" size="+6">【解鎖ＩＰ】</font></h1>
</div>
<hr noshade size="1" color=009900>
<b> 操作完成 </b> <br>
該 IP 已經被解鎖，可以正常登錄。你剛纔的操作<%if logok="0" then Response.Write "<font color=red>沒有</font>"%>被記錄在公開欄中。
<hr noshade size="1" color=009900>
<br>
<table border="1" cellspacing="0" cellpadding="10" bgcolor="E0E0E0" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center">
<form>
<tr>
<td>
<p><%=sj%></p>
<p><%=nickname&"("&userip&")"%></p>
<p><%=unlocklog%></p>
<div align="center">
<input type="button" value="返回" onclick="javascript:history.go(-2)">
</div>
</td>
</tr>
</form>
</table>
<br>
</body>
</html>

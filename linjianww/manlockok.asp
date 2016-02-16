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
nickname=info(0)
if info(2)<10   then Response.Redirect "../error.asp?id=900"
inthechat=Session("ljjh_inthechat")
userip=Request.ServerVariables("REMOTE_ADDR")
if inthechat<>"1"  then Response.Redirect "manerr.asp?id=211"
lockname=Server.HTMLEncode(Trim(Request.Form("lockname")))
lockip=Server.HTMLEncode(Trim(Request.Form("lockip")))
if lockip="" then Response.Redirect "manerr.asp?id=212"
lockwhy=Server.HTMLEncode(Trim(Request.Form("lockwhy")))
if lockwhy="" then Response.Redirect "manerr.asp?id=214"
if CStr(lockname)=CStr(nickname) then Response.Redirect "manerr.asp?id=213"
if len(lockwhy)>60 then lockwhy=Left(lockwhy,60)
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
sql="SELECT ip FROM iplocktemp WHERE ip='" & lockip & "'"
rs.open sql,conn,1,1
if Not(rs.Eof and rs.Bof) then
rs.close
conn.close
set rs=nothing
set conn=nothing
Response.Redirect "manerr.asp?id=215"
end if
rs.close
if lockname<>"" then
Application.Lock
onlinelist=application("ljjh_onlinelist"&session("nowinroom"))
dim newonlinelist()
useronlinename=""
onliners=0
js=1
for i=1 to UBound(onlinelist) step 6
if CStr(onlinelist(i+1))<>CStr(lockname) then
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
if kickip=lockip then
if onliners=0 then
dim listnull(0)
application("ljjh_onlinelist"&session("nowinroom"))=listnull
else
application("ljjh_onlinelist"&session("nowinroom"))=newonlinelist
end if
Application("ljjh_useronlinename"&session("nowinroom"))=useronlinename
Application("ljjh_chatrs"&session("nowinroom"))=onliners
else
Application.UnLock
conn.close
set rs=nothing
set conn=nothing
Response.Redirect "manerr.asp?id=219"
end if
Application.UnLock
sql="SELECT lastkick FROM 用戶 WHERE  姓名='" & lockname & "'"
rs.open sql,conn,1,3
if Not(rs.Eof and rs.Bof) then
rs("lastkick")=sj
rs.Update
end if
rs.close
end if
set rs=nothing
Function SqlStr(data)
SqlStr="'" & Replace(data,"'","''") & "'"
End Function
sql="INSERT INTO iplocktemp (ip,lockdate,locker) VALUES ("
sql=sql & SqlStr(lockip) & ","
sql=sql & SqlStr(sj) & ","
sql=sql & SqlStr(nickname) & ")"
conn.Execute sql
locklog="封鎖IP：<font color=red>" & lockip & "</font>(<font color=FFD7F4>" & lockname & "</font>)！<br><font color=FFD7F4>【原因：" & lockwhy & "】</font>"
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
'sd(199)="<font color=FFD7F4>【封鎖】</font><font color=FFD7F4><font color=red>" & nickname & "</font>封鎖IP：<font color=red>" & lockip & "</font>(" & lockname & ")   原因：" & lockwhy & " </font><font class=t></font>"
Application("ljjh_sd")=sd
Application.UnLock%><html>
<head>
<title>封鎖ＩＰ</title>
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
  <h1><font color="#FF0000" size="+6">【封鎖ＩＰ】</font></h1>
</div>
<hr noshade size="1" color=009900>
<b> 操作完成 </b> <br>
被封鎖的 IP 在 <font color=red>50000</font> 分鐘內不能登錄。你剛纔的操作沒有被記錄在公開欄中。
<hr noshade size="1" color=009900>
<br>
<table border="1" cellspacing="0" cellpadding="10" bgcolor="E0E0E0" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center">
<form>
<tr>
<td>
<p><%=sj%></p>
<p><%=nickname&"("&userip&")"%></p>
<p><%=locklog%></p>
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
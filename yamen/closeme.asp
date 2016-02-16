<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
session("ljjh_jm")=session("ljjh_jm")+1
if session("ljjh_jm")>30 then Response.Redirect "../chat/readonly/bomb.htm"
regjm=Request.form("regjm")
regjm1=Request.form("regjm1")
inroom=trim(Request.form("chat"))
if regjm<>regjm1 then
%>
<script language=vbscript>
alert ("你輸入的認證碼不正確，應該輸入:<%=regjm%>")
location.href = "javascript:history.back()"
</script>
<%
response.end
end if
name=request.form("name")
pass=trim(request.form("pass"))
if name="" or pass="" then
%>
<script language=vbscript>
alert "是不是想開玩了？連大名和口令都不報？"
location.href = "javascript:history.back()"
</script>
<%
Response.End
end if
ljjh_roominfo=split(Application("ljjh_room"),";")
chatroomnum=ubound(ljjh_roominfo)-1
for i=0 to chatroomnum	
	ydl=1
	if Instr(LCase(Application("ljjh_useronlinename"&i))," "&LCase(name)&" ")=0 then ydl=0
	if ydl=1 then exit for
next 
if ydl=0 then
%>
<script language=vbscript>
alert "你在搞什麼呀？你並沒有卡在江湖裡！看看是不是選擇錯了！"
location.href = "javascript:history.back()"
</script>
<%
Response.End
end if
inroom=i
temppass=StrReverse(left(pass&"godxtll,./",10))
templen=len(pass)
mmpassword=""
for j=1 to 10
mmpassword=mmpassword+chr(asc(mid(temppass,j,1))-templen+int(j*1.1))
next
pass=replace(mmpassword,"'","B")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="select 密碼 from 用戶 where 姓名='" & name & "' and 密碼='" & pass & "' "
set rs=conn.execute(sql)
if rs.eof or rs.bof then
%>
<script language=vbscript>
alert "有沒有搞錯？哪有這個人啊？"
location.href = "javascript:history.back()"
</script>
<%
Response.End
end if
if trim(pass)<>rs("密碼") then%>
<script language=vbscript>
alert "密碼不對啊，會不會記錯了？"
location.href = "javascript:history.back()"
</script>
<%
Response.End
end if
%>
<script language=vbscript>
alert "OK,你已經通過了我們的驗證！"
</script>
<%
nickname=name
userip=Request.ServerVariables("REMOTE_ADDR")
kickname=name
kickwhy="我不小線掉線了，沒辦法！"
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
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="對"
sd(199)="<font color=FFD7F4>【掉線管理】</font>"&nickname&"狠狠的揣了自己的小屁股一腳(我掉線了，沒辦法)!" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Application.Lock
onlinelist=Application("ljjh_onlinelist"&inroom)
dim newonlinelist()
useronlinename=""
onliners=0
js=1
for i=1 to UBound(onlinelist) step 6
if CStr(onlinelist(i+1))<>CStr(kickname) then
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
if kickip<>"" then
if onliners=0 then
dim listnull(0)
Application("ljjh_onlinelist"&inroom)=listnull
else
Application("ljjh_onlinelist"&inroom)=newonlinelist
end if
Application("ljjh_useronlinename"&inroom)=useronlinename
Application("ljjh_chatrs"&inroom)=onliners
else
Application.UnLock
end if
%>
<html>
<head>
<title>掉線管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<LINK href="../css.css" rel=stylesheet>
</head>
<body class=p150 background="../images/8.jpg">
<div align="center">
  <table width="360" border="1" cellpadding="10" cellspacing="13" background="../images/12.jpg">
    <tr bgcolor="#FFFFFF">
      <td background="../images/7.jpg"> 
        <div align="center"><font color="#FF6600">【 掉 線 管 理 】</font> </div>
<b><br>
 操作完成 </b> <br>
現在你要以點返回進行登陸了，下次一定要注意使用！一般10分鐘不說話會掉線的，要注意！ <br>
<br>
<table border="1" cellspacing="0" cellpadding="10" bordercolorlight="ff6600" bordercolordark="#FFFFFF" align="center">
<form>
<tr>
<td> <%=sj%><br>
<%=nickname&"("&userip&")"%><br>
現在你已經可以使用了！
<div align="center">
<input type="button" value="關閉窗口" onClick="window.close()"
name="button">
</div>
</td>
</tr>
</form>
</table>

</td>
</tr>
</table>

</div>
</body>
</html> 
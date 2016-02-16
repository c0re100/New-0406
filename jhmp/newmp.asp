<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if InStr(request.form("subject"),"oR")<>0 or InStr(request.form("subject"),"副站長")<>0 or InStr(request.form("subject"),"站長")<>0 or InStr(request.form("subject"),"府")<>0 or InStr(request.form("subject"),"官")<>0 or InStr(request.form("subject"),"江湖總站")<>0 or InStr(request.form("subject"),"官府")<>0 or InStr(request.form("subject"),"Or")<>0 or inStr(request.form("subject"),"&")<>0 or inStr(request.form("subject"),"&&")<>0 or InStr(request.form("subject"),"OR")<>0 or InStr(request.form("subject"),"or")<>0 or InStr(request.form("subject"),"=")<>0 or InStr(request.form("subject"),"`")<>0 or InStr(request.form("subject"),"'")<>0 or InStr(request.form("subject")," ")<>0 or InStr(request.form("subject")," ")<>0 or InStr(request.form("subject"),"'")<>0 or InStr(request.form("subject"),chr(34))<>0 or InStr(request.form("subject"),"\")<>0 or InStr(request.form("subject"),",")<>0 or InStr(request.form("subject"),"<")<>0 or InStr(request.form("subject"),">")<>0 then Response.Redirect "../error.asp?id=54"
if info(3)<30 then Response.Redirect "../error.asp?id=460"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 掌門 from 門派 where 門派='"&trim(request.form("subject"))&"'",conn
if not rs.bof or not rs.eof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('這個門派已經存在了，你不能創建相同的門派！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
rs.close
rs.open "select 等級,銀兩,知質 from 用戶 where id="&info(9),conn
if rs("等級")<100 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('掌門條件戰鬥等級大於100級！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if rs("銀兩")<200000000 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('掌門條件銀兩大於2億！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if rs("知質")<10000 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('掌門條件資質大於10000！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
rs.close
rs.open "select 掌門 from 門派 where 掌門='"&info(0)&"'",conn
if not rs.bof or not rs.eof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你已經是掌門了,回到首頁重進吧！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
str="Insert Into 門派 (門派,簡介,掌門,適合,幫派基金,創立時間) Values('"
str=str & request.form("subject") & "','"
str=str & request.form("comment") & "','"&info(0)&"','"&trim(request.form("ljxb"))&"',100000000,now())"
conn.execute(str)
conn.execute("update 用戶 set 門派='"&trim(request.form("subject"))&"',身份='掌門',grade=5 where id="&info(9))
mp=trim(request.form("subject"))
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: 新細明體}
input {  font-size: 9pt; color: #000000; background-color: #f7f7f7; padding-top: 3px}
.c {  font-family: 新細明體; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: normal; font-variant: normal; text-decoration: none}
--></style>
</head>

<body bgcolor="#000000" text="#000000" link="#0000FF" alink="#0000FF" vlink="#0000FF" background="../linjianww/0064.jpg">
<div align="center">
  <p>&nbsp;</p>
</div>

</body>
</html>
<%
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
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
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="對"
sd(199)="<font color=FFD7F4>【新門派誕生】"& info(0) &"</font>大聲宣布<font color=FFD7F4> "&mp&" </font>新門派正式成立！</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script language=JavaScript>{alert('新門派創建成功！');location.href = 'javascript:history.go(-1)';}</script>"
Response.End

%>
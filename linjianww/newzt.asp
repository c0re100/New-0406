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
if info(2)<10   then Response.Redirect "../error.asp?id=900"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
set rstmp=server.createobject("adodb.recordset")
set rstmp=conn.execute("Select * from 門派 Where 門派='"&trim(request.form("subject"))&"'")
if not rstmp.eof then
Response.Redirect "../error.asp?id=444"
else
str="Insert Into 門派 (門派,簡介,掌門,適合) Values('"
str=str & request.form("subject") & "','"
str=str & request.form("comment") & "','未定','"&trim(request.form("ljxb"))&"')"
conn.execute(str)
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

<body bgcolor="#000000" text="#000000" link="#0000FF" alink="#0000FF" vlink="#0000FF" background="0064.jpg">
<div align="center">
  <p>&nbsp;</p>
</div>

</body>
</html>
<%
rstmp.close
set rs=nothing
set rstmp=nothing
Response.Write "<script language=JavaScript>{alert('新門派添加成功！');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
end if
%>
<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
info=Session("info")
%>
<!--#include file="fun.inc"-->
<%
name1=request.form("name1")
name1=trim(name1)
name1=HTMLEncode(name1)
if instr(name1,"'")<>0 or instr(name1,"|")<>0 or instr(name1," ")<>0 then
response.end
end if
if name1="大家" or name1="無" or instr(name1,"媽")>0 or instr(name1,"爸")>0 or instr(name1,"管理員")>0 or instr(name1,"操")>0 then
response.end
end if
if  name1="" then
message="新名字不能為空"
else
%>
<!--#include file="data.asp"-->
<%
sql="SELECT 名字 FROM 寵物狀態 Where user='" & info(0) & "' and 體力>0"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof Then
message="您沒有寵物或寵物已死，不能改名字!"
else
sql="update 寵物狀態 set 名字='"&name1&"' where user='"&info(0)&"'"
conn.Execute(sql)
message="更換寵物名字成功!"
end if
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>
<html>

<head>
<title>更換名字成功</title>
<style type="text/css"></style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="setup.css">
</head>
<body topmargin="0" leftmargin="0" bgcolor="#3a4b91" text="#000000">
<div align="center"><font color="#000000"><%=message%><br>
<br>
[<a href="feedsheep.asp">返回</a>] </font></div>
</body>
</html>
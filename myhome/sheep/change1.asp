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
if name1="�j�a" or name1="�L" or instr(name1,"��")>0 or instr(name1,"��")>0 or instr(name1,"�޲z��")>0 or instr(name1,"��")>0 then
response.end
end if
if  name1="" then
message="�s�W�r���ର��"
else
%>
<!--#include file="data.asp"-->
<%
sql="SELECT �W�r FROM �d�����A Where user='" & info(0) & "' and ��O>0"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof Then
message="�z�S���d�����d���w���A�����W�r!"
else
sql="update �d�����A set �W�r='"&name1&"' where user='"&info(0)&"'"
conn.Execute(sql)
message="���d���W�r���\!"
end if
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>
<html>

<head>
<title>�󴫦W�r���\</title>
<style type="text/css"></style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="setup.css">
</head>
<body topmargin="0" leftmargin="0" bgcolor="#3a4b91" text="#000000">
<div align="center"><font color="#000000"><%=message%><br>
<br>
[<a href="feedsheep.asp">��^</a>] </font></div>
</body>
</html>
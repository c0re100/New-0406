<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>

<!--#include file="data1.asp"-->
<%sex=request("sex")%>
<html>

<head>
<title>�����d��</title>
</head>
<body>
<table>
<tr><td>�m�W</td><td>�Ы�����</td><td>�ЫΩҦb��</td>

<%sql="select user,house,bigarea from userinfo where sex='"&sex&"'"
set rs=conn.execute(sql)
do while not rs.eof%>


<tr><td><A href="userinfo.asp?id=<%=rs("user")%>"><%=rs("user")%></a></td>
<%if rs("house")="0" then%><td>�S���Ы�</td><td></td>
<%else%>
<td><%=rs("house")%></td>
<td><%=rs("bigarea")%><td>
<%end if%>
</tr>
<%rs.movenext
loop%>
</body>


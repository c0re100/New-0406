<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<!--#include file="jhb.asp"-->
<%
Jname=info(0)
sql="select * from 客戶 where 帳號='" & Jname & "' "
set rs=conn.execute(sql)

%>
<html>
<head>
<title>經紀人辦公室</title>
<link rel="stylesheet" type="text/css" href="style.css">
<style>
td{color:#000000}
p{font-size:16;color:red}
</style>
</head>
<body text="#000000" bgcolor="#993300" >
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
<table><tr><td>
<p align=center style="font-size:14;color:#000000"></p>
</td></tr>
<tr><td>
<p align=center><font color=blue size=2>您的股票帳戶裡現有資金<%=rs("資金")%>兩，您打算取多少？</font></p>
<p align=center style="font-size:14;color:#000000"><form method=POST action='qu1.asp' id=form1 name=form1>
金額：<input type=text name=money  size=10>
<input type=submit value=確定 id=submit1 name=submit1>
</form>
<p align=center><a href=jingji.asp>離開</a></p>
</table></table>
</body>
</html>
<%
set rs=nothing
conn.close
set conn=nothing
%>

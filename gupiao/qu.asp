<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<!--#include file="jhb.asp"-->
<%
Jname=info(0)
sql="select * from �Ȥ� where �b��='" & Jname & "' "
set rs=conn.execute(sql)

%>
<html>
<head>
<title>�g���H�줽��</title>
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
<p align=center><font color=blue size=2>�z���Ѳ��b��̲{�����<%=rs("���")%>��A�z������h�֡H</font></p>
<p align=center style="font-size:14;color:#000000"><form method=POST action='qu1.asp' id=form1 name=form1>
���B�G<input type=text name=money  size=10>
<input type=submit value=�T�w id=submit1 name=submit1>
</form>
<p align=center><a href=jingji.asp>���}</a></p>
</table></table>
</body>
</html>
<%
set rs=nothing
conn.close
set conn=nothing
%>

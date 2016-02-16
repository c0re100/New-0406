<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
jin=request.form("money")
jin=abs(int(jin))
mess=""
Jname=info(0)
sql="select 銀兩 from 用戶 where 姓名='" & Jname & "' "
set rs=conn.execute(sql)
yin=rs("銀兩")
if jin>yin then
mess="你的股票經紀人大聲說，你哪裡這麼多錢，想坑我嗎？"
else
sql="update 用戶 set 銀兩=銀兩-"& jin & " where 姓名='" & Jname & "'"
conn.execute sql
%>
<!--#include file="jhb.asp"-->
<%
sql="update 客戶 set 資金=資金+" & jin & " where 帳號='" & Jname & "'"
conn.execute sql
mess="你已經把"&jin&"兩存進了你的股票帳戶"
end if
%>

<head>
<title>成功</title>
</head>

<body text="#000000" bgcolor="#990000">
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
<table width=100%>
<tr><td>
<p align=center style='font-size:14;color:red'><%=mess%></p>
<p align=center><a href=cun.asp>返回</a></p>
</td></tr>
</table>
</td></tr>
</table>
</body>
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
sql="select �Ȩ� from �Τ� where �m�W='" & Jname & "' "
set rs=conn.execute(sql)
yin=rs("�Ȩ�")
if jin>yin then
mess="�A���Ѳ��g���H�j�n���A�A���̳o��h���A�Q�|�ڶܡH"
else
sql="update �Τ� set �Ȩ�=�Ȩ�-"& jin & " where �m�W='" & Jname & "'"
conn.execute sql
%>
<!--#include file="jhb.asp"-->
<%
sql="update �Ȥ� set ���=���+" & jin & " where �b��='" & Jname & "'"
conn.execute sql
mess="�A�w�g��"&jin&"��s�i�F�A���Ѳ��b��"
end if
%>

<head>
<title>���\</title>
</head>

<body text="#000000" bgcolor="#990000">
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
<table width=100%>
<tr><td>
<p align=center style='font-size:14;color:red'><%=mess%></p>
<p align=center><a href=cun.asp>��^</a></p>
</td></tr>
</table>
</td></tr>
</table>
</body>
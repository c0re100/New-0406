<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<!--#include file="../config.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%>
<!--#include file="data.asp"-->
<%
selectaction=request.form("s1")
selecthouse=request.form("r1")
bigarea=request.form("area")
if selectaction="" or selecthouse="" then
response.redirect"warning.asp"
else

select case selecthouse
case "house1"
sploshtemp=800
case "house2"
sploshtemp=20000
case "house3"
sploshtemp=50000
case "house4"
sploshtemp=200000
case "house5"
sploshtemp=500000
case "house6"
sploshtemp=1000000
end select
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�F�C����-�ЫΥ��</title>
<link rel="StyleSheet" href="new.css" title="Contemporary">
</head>

<body topmargin="0" leftmargin="0" background="../ljimage/11.jpg">
<div align="center">
</div>
<div align="center">
<center>
<table border="0" width="468" bordercolor="#FFFFFF" cellspacing="1"
cellpadding="0" height="1">
<tr>
<td class="p3" width="415" height="30" valign="top">
<%
select case selectaction
case "buy"
set rs=conn.execute("select �Ȩ� from �Τ� where �m�W='"&info(0)&"'")
set rs=conn.execute("select house,bigarea from userinfo where user='"&info(0)&"'")
if rs("house")<>"0" then
%>
<table>
<tr>
<td><font size="2">�z����֦��@�ҩЫΡI�I�p�G�Q�n���ЫΡA�Х��汼�H�e���ЫΡI�I�z���Ы������O<%=rs("house")%>�Ҧb�Ϭ��G<%=rs("bigarea")%></font></td>
</table>
<%
else
set rs=conn.execute("select �Ȩ� from �Τ� where �m�W='"&info(0)&"'")
if rs("�Ȩ�")<sploshtemp then
%>
<table>
<tr>
<td><font size="2">�z�S�������������r�I�I</font></td>
</table>
<%
else
sploshtemp=rs("�Ȩ�")-sploshtemp
rs.close
conn.execute"update �Τ� set �Ȩ�='"&sploshtemp&"' where �m�W='"&info(0)&"'"
conn.execute"update userinfo set house='"&selecthouse&"',bigarea='"&bigarea&"',homeopen='1' where user='"&info(0)&"'"
'set ls=conn.execute("select * from trueinfo where user='"&info(0)&"'")
'if ls.eof then
'conn.execute"insert into trueinfo(user) values ('"&info(0)&"')"
'end if
conn.close
conn1.close
%>
<table>
<tr>
<td><font size="2">���߱z�I�z�w�g���\���ʶR�F�Фl�I�I �G-�^</font></td>
</table>
<%
end if
end if

case "sale"
set rs=conn.execute("select �Ȩ� from �Τ� where �m�W='"&info(0)&"'")
set rs=conn1.execute("select house from userinfo where user='"&info(0)&"'")

if rs("house")<>selecthouse then
%>
<table>
<tr>
<td><font size="2">�z�S���o�˪��ЫΥi��I</font></td>
</table>
<%
else
set rs=conn.execute("select �Ȩ� from �Τ� where �m�W='"&info(0)&"'")
sploshtemp=sploshtemp/4
sploshtemp=rs("�Ȩ�")+sploshtemp
rs.close
conn.execute"update �Τ� set �Ȩ�='"&sploshtemp&"' where �m�W='"&info(0)&"'"
conn.execute"update userinfo set house='0',bigarea=' 'where user='"&info(0)&"'"
conn.close
%>
<table>
<tr>
<td><font size="2">���ߡI�z���\����X�Фl�A����F���1/4���Ȥl�C�o�̯u�¡G�]</font></td>
</table>
<%
end if
end select
end if

%>
</td>
</tr>
<tr>
<td class="p2" width="415" height="74" valign="top"></td>
</tr>
</table>
</center>
</div>

</body>

</html>

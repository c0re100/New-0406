<%
Response.Expires=0
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
name=info(0)
Set Conn=Server.CreateObject("ADODB.Connection")
Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("global.asa")
Conn.Open Connstr
set rs=conn.execute("select * from ������~ where �֦���='"& name &"'")
%>
<html>
<head>
<title>�ӫ~�@�l</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
td {  font-family: "�s�ө���"; font-size: 12px}
input {  font-family: "�s�ө���"; font-size: 12px; border: thin #333333 dotted; color: #660000; background-color: #FFFFCC}
-->
</style>
</head>

<body bgcolor="#FFFFFF" text="#660000" background="../../images/8.jpg">
<br>
<br>
<table width="60%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr align="center"> 
    <td height="22"><big><strong></strong></big>�A�{���ӫ~</td>
  </tr>
</table>
<table width="60%" border="1" cellspacing="4" cellpadding="4" align="center">
  <tr align="center" bgcolor="#660000"> 
    <td width="25%"><font color="#FFFFFF">����</font></td>
    <td width="25%"><font color="#FFFFFF">�H��</font></td>
    <td width="25%"><font color="#FFFFFF">�]�_</font></td>
    <td width="25%"><font color="#FFFFFF">����</font></td>
  </tr>
<% do while not rs.eof %>
  <tr align="center"> 
    <td width="25%"><%=rs("����")%></td>
    <td width="25%"><%=rs("�H��")%></td>
    <td width="25%"><%=rs("�]�_")%></td>
    <td width="25%"><%=rs("����")%></td>
  </tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
<table width="60%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr align="center"> 
    <td height="32"><br>
      �F�C���� <br>
    </td>
  </tr>
</table>
</body>
</html>
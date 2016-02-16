<%
Response.Expires=0
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
name=info(0)
Set Conn=Server.CreateObject("ADODB.Connection")
Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("global.asa")
Conn.Open Connstr
set rs=conn.execute("select * from 航海物品 where 擁有者='"& name &"'")
%>
<html>
<head>
<title>商品一纜</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
td {  font-family: "新細明體"; font-size: 12px}
input {  font-family: "新細明體"; font-size: 12px; border: thin #333333 dotted; color: #660000; background-color: #FFFFCC}
-->
</style>
</head>

<body bgcolor="#FFFFFF" text="#660000" background="../../images/8.jpg">
<br>
<br>
<table width="60%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr align="center"> 
    <td height="22"><big><strong></strong></big>你現有商品</td>
  </tr>
</table>
<table width="60%" border="1" cellspacing="4" cellpadding="4" align="center">
  <tr align="center" bgcolor="#660000"> 
    <td width="25%"><font color="#FFFFFF">絲綢</font></td>
    <td width="25%"><font color="#FFFFFF">人參</font></td>
    <td width="25%"><font color="#FFFFFF">珠寶</font></td>
    <td width="25%"><font color="#FFFFFF">瓷器</font></td>
  </tr>
<% do while not rs.eof %>
  <tr align="center"> 
    <td width="25%"><%=rs("絲綢")%></td>
    <td width="25%"><%=rs("人參")%></td>
    <td width="25%"><%=rs("珠寶")%></td>
    <td width="25%"><%=rs("瓷器")%></td>
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
      靈劍江湖 <br>
    </td>
  </tr>
</table>
</body>
</html>
<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
name=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT 物品名,數量,功效 FROM 修練物品 WHERE 擁有者=" & info(9) & " and 數量>0  order by 增加數值",conn
%>
<html>
<head>
<title>配置藥材</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="style.css">
</head>
<body bgcolor="#660000" text="#FFFFFF" leftmargin="0" topmargin="0" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<div align="left">
  <div align="center">現已配藥材一覽 <font face="幼圓"><a href="javascript:this.location.reload()">刷新</a></font></div>
<div align="center"> <br>
<table border="1" align="center" width="141" cellpadding="1" cellspacing="0" height="31">
<tr align="center">
<td nowrap width="64">
<div align="center"><font color="#FFFFFF" size="-1">物品名</font></div>
</td>
<td nowrap width="33">
<div align="center"><font color="#FFFFFF" size="-1">數量 </font> </div>
<td colspan="2" nowrap width="36">
<div align="center"><font color="#FFFFFF" size="-1">功效</font></div>
</td>
</tr>
<%
do while not rs.eof
%>
<tr>
<td width="64" height="3">
<div align="center"><font color="#FFFFFF" size="-1"><%=rs("物品名")%>
</font> </div>
</td>
<td width="33" height="3">
<div align="center"><font color="#FFFFFF" size="-1"><%=rs("數量")%>
</font> </div>
<td colspan="2" height="3" width="36">
<div align="center"><font color="#FFFFFF" size="-1"><%=rs("功效")%>
</font> </div>
</td>
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
</div>
</div>
</body>
</html>

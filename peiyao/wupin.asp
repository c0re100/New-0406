<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(0)="" then Response.Redirect "../error.asp?id=210"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
Response.Expires=0
name=info(0)
rs.open "SELECT 物品名,數量,銀兩 FROM 物品 WHERE 擁有者=" & info(9) & " and 數量>0  and (物品名='大鯊魚' or 物品名='冰水' or 物品名='礦石' or 物品名='大草魚' or 物品名='老虎肉' or 物品名='小鯉魚') order by 銀兩 ",conn
%>
<html>
<head>
<title>藥材</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="../style.css">
</head>
<body bgcolor="#660000" text="#FFFFFF" leftmargin="0" topmargin="0">
<div align="left">
  <div align="center"><font color="#FFFFFF">現有藥材一覽 </font><font face="幼圓"><a href="javascript:this.location.reload()"><font color="#FFFF00">刷新</font></a></font></div>
<div align="center"> <br>
<table border="1" align="center" width="141" cellpadding="1" cellspacing="0" height="31">
<tr align="center">
<td nowrap width="64" height="16">
<div align="center"><font color="#FFFFFF" size="-1">物品名</font></div>
</td>
<td nowrap width="33" height="16">
<div align="center"><font color="#FFFFFF" size="-1">數量 </font> </div>
<td colspan="2" nowrap width="36" height="16">
<div align="center"><font color="#FFFFFF" size="-1">價值</font></div>
</td>
</tr>
<%
do while not rs.eof
%>
<tr>
<form method=POST action='wupin1.asp?action=<%=rs("物品名")%>&name=<%=name%>'>
<td width="64" height="3">
<div align="center"><font color="#FFFFFF" size="-1"><%=rs("物品名")%>
</font> </div>
</td>
<td width="33" height="3">
<div align="center"><font color="#FFFFFF" size="-1"><%=rs("數量")%>
</font> </div>
<td colspan="2" height="3" width="36">
<div align="center"><font color="#FFFFFF" size="-1"><%=rs("銀兩")%>
</font> </div>
</td>
</form>
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

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
<%
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"%>
<html>
<head>
<title>兵器鋪</title>
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: 新細明體}
.c {  font-family: 新細明體; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
<style>
td{color:#000000}
p{font-size:16;color:red}
</style>
</head>
<body bgcolor=#000000 text="#FFFFFF" background="0064.jpg">
<p align=center><font size="+6">兵器庫管理</font></p>
<br>
<table border=0 bgcolor="#BEE0FC" align=center width=550 cellpadding="10" cellspacing="2">
<tr bgcolor="#336633">
<td height="17" colspan="4">
<p align=center>現 有 暗 器</P>
<tr bgcolor=#C4DEFF>
<td height="18">
<div align="center"><font size="2">暗器名</font></div>
</td>
<td bgcolor="#C4DEFF" height="18">
<div align="center"><font size="2">功 能</font></div>
</td>
<td height="18">
<div align="center"><font size="2">售 價</font></div>
</td>
<td height="18">
<div align="center"><font size="2">操 作</font></div>
</td>
</tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="SELECT * FROM 物品買 where  類型='暗器'"
Set Rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
<tr bgcolor=#f7f7f7>
<td height="11"><font size="2"><%=rs("物品名")%> </font>
<td height="11"><font size="2">傷內力<%=abs(rs("內力"))%>，傷體力<%=abs(rs("體力"))%></font></td>
<td height="11"><font size="2"><%=rs("銀兩")%>兩</font></td>
<td height="11"><a href="modifyyao.asp?wupinid=<%=rs("id")%>">修改</a></td>
</tr>
<%
rs.movenext
loop
%>
</table>
<p align=center></P>
<table width="551" border="0" cellspacing="1" cellpadding="2" align="center" bgcolor="#BEE0FC">
<tr bgcolor="#336633">
<td colspan="5" height="32">
<p align=center>現 有 裝 備</p>
</td>
</tr>
<tr bgcolor="#C4DEFF">
<td width="103" height="20">
<div align="center"><font size="2">裝備名</font></div>
</td>
<td width="60" height="20">
<div align="center"><font size="-1">類型</font></div>
</td>
<td width="193" height="20"><font size="2">功 能</font></td>
<td width="98" height="20">
<div align="center"><font size="2">售 價</font></div>
</td>
<td width="71" height="20">
<div align="center"><font size="2">操 作</font></div>
</td>
</tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="SELECT * FROM 物品買 where  類型='左手' or 類型='右手' or 類型='雙腳' or 類型='頭盔' or 類型='盔甲' or 類型='裝飾' or 類型='裝備'"
Set Rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
<tr>
<td bgcolor="#FFFFFF" width="103" height="23"><font size="2"><%=rs("物品名")%></font></td>
<td bgcolor="#FFFFFF" width="60" height="23"><font size="2"><%=rs("類型")%>
</font></td>
<td bgcolor="#FFFFFF" width="193" height="23"><font size="2">加防御<%=abs(rs("防御"))%>，加攻擊<%=abs(rs("攻擊"))%></font></td>
<td bgcolor="#FFFFFF" width="98" height="23"><font size="2"><%=rs("銀兩")%>兩</font></td>
<td bgcolor="#FFFFFF" width="71" height="23"><a href="modifyyao.asp?wupinid=<%=rs("id")%>">修改</a></td>
</tr>
<%
rs.movenext
loop
conn.close
%>
</table>
</body>
</html>
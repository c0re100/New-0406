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
<html>
<head>
<title>門派管理</title>
<style type="text/css">
<!--
p            { line-height: 20px; font-size: 9pt }
table        { font-size: 9pt }
a:link       { color: #FF9900; text-decoration: none }
-->
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body text="#000000" vlink="#FF9900" topmargin="0"
leftmargin="0" bgcolor="#660000">
<p align="center"> 
  <%
id=request("id")
my=request("my")
%>
  <font color="#CC0000" face="幼圓"><a href="javascript:this.location.reload()"><font color="#FFFF00">刷新</font></a></font> 
<table border="0" width="500" cellspacing="0" cellpadding="0"
background="bg.gif" align="center">
<tr align="center">
<td background="top1.gif" width="500" height="26"><font
color="#FF6600"><b><font size="+1">江 湖 門 派</font></b></font></td>
</tr>
<tr align="center">
<td>
<table width="470" border="1" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" height="13">
<tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT id,姓名,性別,身份,regip,regtime,lasttime,times,門派 FROM 用戶 where 門派='"&id&"'",conn
%>
<td width="31">
<div align="center"><font color="#FFFFFF">ID</font></div>
</td>
<td width="54">
<div align="center"><font color="#FFFFFF">姓名</font></div>
</td>
<td width="27">
<div align="center"><font color="#FFFFFF">性別</font></div>
</td>
<td width="74">
<div align="center"><font color="#FFFFFF">身份</font></div>
</td>
<td width="66">
<div align="center"><font color="#FFFFFF">注冊ip</font></div>
</td>
<td width="65">
<div align="center"><font color="#FFFFFF">注冊時間</font></div>
</td>
<td width="32">
<div align="center"><font color="#FFFFFF">登錄</font></div>
</td>
<td width="73">
<div align="center"><font color="#FFFFFF">掌門操作</font></div>
</td>
<td width="28">
<div align="center"><font color="#FFFFFF">刪除</font></div>
</td>
</tr>
<%
do while not rs.bof and not rs.eof
%>
<tr>
<td width="31" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("ID")%></font></div>
</td>
<td width="54" height="30">
<div align="center"><a href="../yamen/mt.asp?action=<%=rs("姓名")%>"><font color="#FFFFFF"><%=rs("姓名")%></font></a></div>
</td>
<td width="27" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("性別")%></font></div>
</td>
<td width="74" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("身份")%></font></div>
</td>
<td width="66" height="30">
<div align="center"><font color="#FFFFFF">
<%if info(2)>=9  then%>
	<%=rs("regip")%>
<%else%>
	無權查看
<%end if%>
</font></div>
</td>
<td width="65" height="30">
<div align="center"><font color="#FFFFFF">
<%if info(2)>=9  then%>
	<%=rs("regtime")%><br>
	<%=rs("lasttime")%>
<%else%>
	<%=rs("regtime")%>
<%end if%>
</font></div>
</td>
<td width="32" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("times")%></font></div>
</td>
<td width="73" height="30">
<div align="center"><a href="mp3.asp?you=<%=rs("姓名")%>&amp;id=<%=rs("門派")%>"><font color="#FFFFFF">逐出師門</font></a></div>
</td>
<td width="28" height="30">
<div align="center"><a href="mp4.asp?you=<%=rs("姓名")%>&id=<%=rs("門派")%>"><font color="#FFFFFF">
<%if info(2)>=9 then%>
	刪除
<%else%>
	無權
<%end if%>
</font></a></div>
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
</td>
</tr>
</table>
</body>
</html>


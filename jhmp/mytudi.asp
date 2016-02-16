<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=n & "-" & y & "-" & r & " " & s & ":" & f & ":" & m
%>
<html>
<head>
<title>個人所收徒弟</title>
<style type="text/css">
<!--
p            { line-height: 20px; font-size: 9pt }
table        { font-size: 9pt }
a:link       { color: #FF9900; text-decoration: none }
-->
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../chat/ccss2.css" rel="stylesheet" type="text/css">
</head>

<body vlink="#0000CC" topmargin="0"
leftmargin="0" bgcolor="#990000" alink="#0000CC" link="#3300FF" background="../ff.jpg">
<table width="92%" border="0" cellpadding="2" cellspacing="2" height="13" align="center" bordercolor="#000000">
  <tr bgcolor="#CC0000"> 
    <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT id,姓名,性別,門派,身份,lasttime,mvalue,grade,等級 FROM 用戶 where  師傅='"& info(0)&"'",conn
%>
    <td width="37" height="10"> <div align="center"><font color="#FFFFFF">ID</font></div></td>
    <td width="87" height="10"> <div align="center"><font color="#FFFFFF">姓名</font></div></td>
    <td width="37" height="10"> <div align="center"><font color="#FFFFFF">性別</font></div></td>
    <td width="74" height="10"> <div align="center"><font color="#FFFFFF">門派</font></div></td>
    <td width="82" height="10"> <div align="center"><font color="#FFFFFF">身份</font></div></td>
    <td width="161" height="10"> <div align="center"><font color="#FFFFFF">最後登陸</font></div></td>
    <td width="68" height="10"> <div align="center"><font color="#FFFFFF">月泡點</font></div></td>
    <td width="50" height="10"> <div align="center"><font color="#FFFFFF">管理級</font></div></td>
    <td width="79" height="10"> <div align="center"><font color="#FFFFFF">戰鬥級</font></div></td>
  </tr>
  <%
jl=0
do while not rs.bof and not rs.eof
jl=jl+1
%>
  <tr bgcolor="#FFCC00"> 
    <td width="37" height="30"> <div align="center"><font color="#000000"><%=rs("ID")%></font></div></td>
    <td width="87" height="30"> <div align="center"><font color="#0000FF"><%=rs("姓名")%></font></div></td>
    <td width="37" height="30"> <div align="center"><font color="#000000"><%=rs("性別")%></font></div></td>
    <td width="74" height="30"> <div align="center"><font color="#000000"><%=rs("門派")%></font></div></td>
    <td width="82" height="30"> <div align="center"><font color="#000000"><%=rs("身份")%></font></div></td>
    <td width="161" height="30"> <div align="center"><font color="#000000"><%=rs("lasttime")%></font></div></td>
    <td width="68" height="30"> <div align="center"><font color="#FF0000"><%=rs("mvalue")%></font></div></td>
    <td width="50" height="30"> <div align="center"><font color="#000000"><%=rs("grade")%></font></div></td>
    <td width="79" height="30"> <div align="center"><font color="#000000"><%=rs("等級")%></font></div></td>
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
<p align="center"> <font color="#000000"><br>
  </font><font size="+3">徒弟總數:<b><%=(jl)%></b>人</font> 
<div align="center">
  <p>&nbsp;</p>
 
</div>
</body>
</html>
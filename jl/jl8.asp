<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
dim page
page=request.querystring("page")
PageSize = 15
rs.open "delete * from 操作記錄 where 時間<now()-3",conn,3,3
rs.open "Select 姓名1,時間,姓名2,數據 From 操作記錄 where 方式='冊封' Order by 時間 DESC",conn,3,3
rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page
%>
<html>

<head>
<title>江湖冊封</title>
<style type="text/css">td           { font-family: 新細明體; font-size: 9pt }
body         { font-family: 新細明體; font-size: 9pt }
select       { font-family: 新細明體; font-size: 9pt }
a            { color: #FFC106; font-family: 新細明體; font-size: 9pt; text-decoration: none }
a:hover      { color: #cc0033; font-family: 新細明體; font-size: 9pt; text-decoration:
underline }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body leftmargin="0" topmargin="0" background="../images/8.jpg">
<table border="0" height="24" width="91%" cellspacing="0" cellpadding="0"
align="center">
<tbody>
<tr>
<td height="15" width="100%" bgcolor="#669999"><font color="#669966"> <font color="#FFFFFF"><b>江湖冊封</b></font></font></td>
</tr>
</tbody>
</table>
<div align="center">
<table width="91%" align="center" cellspacing="0" border="0"
cellpadding="0">
<tr>
<td width="100%">
<table border="1" cellspacing="1" cellpadding="0" width="97%"
bordercolorlight="#EFEFEF">
<tr bgcolor="#FFFFFF">
<td width="11%">
<div align="center"><font color="#FF6600"> 操 作 者 </font></div>
</td>
<td width="18%">
<div align="center"><font color="#FF6600"> 時 間 </font></div>
</td>
<td width="12%">
<div align="center"><font color="#FF6600"> 被操作者 </font></div>
</td>
<td width="59%">
<div align="center"><font color="#FF6600"> 事 件 數 據</font></div>
</td>
</tr>
<%
count=0
do while not (rs.eof or rs.bof) and count<rs.PageSize
%>
<tr bgcolor="#FFCC99"  onMouseOut="this.bgColor='#FFCC99';"onMouseOver="this.bgColor='#DFEFFF';"> 
<td width="11%">
<div align="center"> <font color="#0000FF"><%=rs("姓名1")%></font>
</div>
</td>
<td width="18%">
<div align="center"> <%=rs("時間")%> </div>
</td>
<td width="12%">
<div align="center"> <%=rs("姓名2")%> </div>
</td>
<td width="59%">
<div align="center"> <%=Replace(rs("數據"),"red","blue")%> </div>
</td>
</tr>
<%rs.movenext%>
<%count=count+1%>
<%loop
pa=rs.pagecount
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
<table border="0" cellspacing="1" cellpadding="1" width="100%" bordercolorlight="#EFEFEF">
<tr>
<td align="left" width="37%" height="2">[共<font color="red"><b><%=pa%></b></font>頁]
</td>
<td align="right" width="63%" height="2">
<div align="center">[<a href="jl8.asp?page=<%=page-1%>">上一頁</a>][第<%=page%>頁][<a href="jl8.asp?page=<%=page+1%>">下一頁</a>]</div>
</td>
</tr>
</table>
</table>
</div>

</body>
<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
dim page
page=request.querystring("page")
PageSize = 15
myname=request.cookies("Jname")
myname=info(0)
rs.open "SELECT * FROM 管理記錄 order by 時間 DESC",conn,3,3
rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page
%>
<html>
<head>

<title>記錄管理</title>
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: 新細明體}
.c {  font-family: 新細明體; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
</head>

<body background="0064.jpg">
<div align="center"> 
  <p><br>
    <font color="#009900"><b><font color="#FF0000" size="+6">[記錄管理]</font></b></font><font
color="#FFFFFF"><br>
    </font> </p>
  </div>
<table cellpadding="1" cellspacing="0" border="1" align="center" width="700"
bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr bgcolor="#99CCFF"> 
    <td height="27" width="148"> 
      <div align="center"><font color="#FF6600">姓名</font></div>
    </td>
    <td height="27" width="166"> 
      <div align="center"><font color="#FF6600">時間</font></div>
    </td>
    <td height="27" width="156"> 
      <div align="center"><font color="#FF6600">IP</font></div>
    </td>
    <td height="27" width="212"> 
      <div align="center"><font color="#FF6600">方式</font></div>
    </td>
  </tr>
  <form method=POST action='fawupinok.asp'>
    <%
count=0
do while not rs.eof and count<rs.PageSize
%>
	<tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td height="2" width="148"> 
        <div align="center"> <%=rs("姓名")%> </div>
      </td>
      <td height="2" width="166"> 
        <div align="center"> <%=rs("時間")%></div>
      </td>
      <td height="2" width="156"><%=rs("ip")%></td>
      <td height="2" width="212"> 
        <div align="center"> <%=rs("記錄")%></div>
      </td>
    </tr>
	<%rs.movenext%>
<%count=count+1%>
<%loop
%>
  </form>
</table><tr bgcolor="#f7f7f7">
    <td align="left">
    <div align="right"><font size="2">[共<font color="#990000"><b><%=rs.pagecount%></b></font>頁][<a
href="managelist.asp?page=<%=page-1%>"><font color="#990000">上一頁</font></a>][第<%=page%>頁][<a
href="managelist.asp?page=<%=page+1%>"><font color="#990000">下一頁</font></a>]</font></div>
  </td>
</tr><%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</body>

</html>
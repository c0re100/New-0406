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
%>
<html>
<head>

<title>管理記錄追蹤</title>
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
    <font color="#009900"><b><font color="#FF0000" size="+6">[ 管理記錄追蹤]</font></b></font><font
color="#FFFFFF"><br>
    </font> </p>
  </div>
<table cellpadding="1" cellspacing="0" border="1" align="center" width="700"
bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr bgcolor="#99CCFF"> 
    <td height="27" width="77"> 
      <div align="center"><font color="#FF6600">官府人員</font></div>
    </td>
    <td height="27" width="114"> 
      <div align="center"><font color="#FF6600">操作時間</font></div>
    </td>
    <td height="27" width="142"> 
      <div align="center"><font color="#FF6600"> IP記錄</font></div>
    </td>
    <td height="27" width="142" bgcolor="#99CCFF"> 
      <div align="center"><font color="#FF6600"> 操作過程</font></div>
    </td>
    <td height="27" width="150"> 
      <div align="center"><font color="#FF6600"> 主 站 長 操 作 </font></div>
    </td>
  </tr>
  <%Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("ljjh_usermdb")
rs.open "SELECT * FROM 操作記錄 where 方式='管理操作'  order by id",conn
s=0
do while not rs.eof and not rs.bof
s=s+1
%>
  <form method=POST action='ljyamenok.asp?a=m&id=<%=rs("id")%>'>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td height="2" width="77" bgcolor="#FFFFFF"><font color="#FF6600"><%=rs("姓名1")%></font></td>
      <td height="2" width="114" bgcolor="#FFFFFF"><font color="#FF6600"><%=rs("時間")%></font></td>
      <td height="2" width="142"> 
        <div align="center"> <font color="#FF6600"><%=rs("姓名2")%></font></div>
      </td>
      <td height="2" width="142"> 
        <div align="center"> <font color="#FF6600"><%=rs("數據")%></font></div>
      </td>
      <td height="2" width="150"> 
        <div align="center"> <a href="ljyamenok.asp?a=m&id=<%=rs("id")%>"> </a> 
          <input type="submit" value="刪除"
name="submit" >
        </div>
      </td>
    </tr>
  </form>
  <%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table></body></html>
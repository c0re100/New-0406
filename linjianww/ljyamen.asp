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

<title>官府人員管理</title>
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
    <font color="#009900"><b><font color="#FF0000" size="+6">[ 官府人員管理]</font></b></font><font
color="#FFFFFF"><br>
    </font> </p>
  </div>
<table cellpadding="1" cellspacing="0" border="1" align="center" width="700"
bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr bgcolor="#99CCFF"> 
    <td height="27" width="88"> 
      <div align="center"><font color="#FF6600">注冊ID</font></div>
    </td>
    <td height="27" width="77"> 
      <div align="center"><font color="#FF6600">姓 名</font></div>
    </td>
    <td height="27" width="114"> 
      <div align="center"><font color="#FF6600">最後登陸</font></div>
    </td>
    <td height="27" width="142" bgcolor="#99CCFF"> 
      <div align="center"><font color="#FF6600"> 身 份</font></div>
    </td>
    <td height="27" width="103"> 
      <div align="center"><font color="#FF6600"> 管理等級</font></div>
    </td>
    <td height="27" width="150"> 
      <div align="center"><font color="#FF6600"> 站 長 操 作 </font></div>
    </td>
  </tr>
  <%Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("ljjh_usermdb")
rs.open "SELECT * FROM 用戶 where 門派='官府' or grade>=6 order by id",conn
s=0
do while not rs.eof and not rs.bof
s=s+1
%>
  <form method=POST action='ljyamenok.asp?a=m&id=<%=rs("id")%>'>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td height="2" width="88"> 
        <div align="center"><font color="#FF6600">ID:</font> <font color="#FF6600"><%=rs("ID")%></font></div>
        <div align="center"></div>
      </td>
      <td height="2" width="77" bgcolor="#FFFFFF"><font color="#FF6600"><%=rs("姓名")%></font></td>
      <td height="2" width="114" bgcolor="#FFFFFF"><font color="#FF6600"><%=rs("lasttime")%></font></td>
      <td height="2" width="142"> 
        <div align="center"> 
          <input type="text" name="shenfen" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="20"
value="<%=rs("身份")%>" maxlength="10">
        </div>
      </td>
      <td height="2" width="103"> 
        <div align="center"> 
          <input type="text" name="dj" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10"
value="<%=rs("grade")%>" maxlength="4">
        </div>
      </td>
      <td height="2" width="150"> 
        <div align="center"> <a href="ljyamenok.asp?a=m&id=<%=rs("id")%>"> 
          <input type="submit" value="修改"
name="submit">
          &nbsp;&nbsp; </a> 
          <input type="submit" value="開除"
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
if s<50 then
%>
  <form method=POST action='ljyamenok.asp?a=n'>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td width="88" height="2"> 
        <div align="center"><font color="#FF6600">添加官府人員</font></div>
      </td>
      <td height="2" colspan="2"> 
        <input type="text" name="name1" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="20"
maxlength="10">
      </td>
      <td width="142" height="2"> 
        <div align="center"> 
          <input type="text" name="shenfen1" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="20"
maxlength="10">
        </div>
      </td>
      <td width="103" height="2"> 
        <div align="center"> 
          <input type="text" name="dj1" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10"
maxlength="4">
        </div>
      </td>
      <td width="150" height="2"> 
        <div align="center"> 
          <input type="submit" value="添加"
name="submit">
        </div>
      </td>
    </tr>
  </form>
  <%end if%>
</table>
</body>
</html>
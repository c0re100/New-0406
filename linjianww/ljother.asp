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

<title>雜項管理</title>
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: 新細明體}
.c {  font-family: 新細明體; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
</head>

<body background="0064.jpg">
<table cellpadding="1" cellspacing="0" border="1" align="center" width="670"
bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr bgcolor="#99CCFF"> 
    <td height="27"> 
      <div align="center"><b><font color="#FF0000" size="+6">[雜項管理]</font></b></div>
    </td>
  </tr>
  <%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "Select * from 集合 where id>=0 order by id",conn
do while not rs.eof and not rs.bof
%>
  <form method=POST action='ljotherok.asp'>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td height="2">挖寶開關 
        <input type="text" name="wa" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="2"
value="<%=rs("wabao")%>" maxlength="2">
        1--------- 挖寶僅對會員開放 0-----------挖寶對所有人開放</td>
    </tr>
    <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
      <td height="2"> 
        <p> 會員表查看 
          <input type="text" name="hy" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="2"
value="<%=rs("hysee")%>" maxlength="2">
          1------- 會員列表觀看僅對會員開放 0---------會員列表觀看對所有人開放</p>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
      <td height="2"> 
        <p> 煉寶開關
<input type="text" name="lianbao" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="2"
value="<%=rs("lianbao")%>" maxlength="2">
          1------- 煉寶僅對會員開放 0---------煉寶對所有人開放</p>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td height="2"> 
        <div align="center"> 
          <input type="submit" value="修改更新"
name="submit">
        </div>
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
  </form>
</table>
<p class="p1" align="center">以上空格處不能為空<br>
</p>
</body>
</html>
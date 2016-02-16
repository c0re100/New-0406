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

<title>贈送管理</title>
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
    <font color="#009900"><b><font color="#FF0000" size="+6">[贈送管理]</font></b></font><font
color="#FFFFFF"><br>
    </font> </p>
  </div>
<table cellpadding="1" cellspacing="0" border="1" align="center" width="700"
bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr bgcolor="#99CCFF"> 
    <td height="27" width="77"> 
      <div align="center"><font color="#FF6600">姓名</font></div>
    </td>
    <td height="27" width="114"> 
      <div align="center"><font color="#FF6600">流星</font></div>
    </td>
    <td height="27" width="114"> 
      <div align="center"><font color="#FF6600">流星淚</font></div>
    </td>
    <td height="27" width="250"> 
      <div align="center"><font color="#FF6600">操作</font></div>
    </td>
  </tr>
    <form method=POST action='sendwupinok.asp'>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td height="2" width="142"> 
        <div align="center"> 
          <input type="text" name="a" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="12" maxlength="12">
        </div>
      </td>
      <td height="2" width="142"> 
        <div align="center"> 
          <input type="text" name="b" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="4" maxlength="4" value="0">
        </div>
      </td>
      <td height="2" width="142"> 
        <div align="center"> 
          <input type="text" name="c" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="4" maxlength="4" value="0">
        </div>
      </td>
      <td height="2" width="150"> 
        <div align="center"> 
          <input type="submit" value="發放"
name="submit">
          &nbsp;&nbsp; </div>
      </td>
    </tr>
  </form>

</table>
</body>

</html>
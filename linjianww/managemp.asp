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
id=Request.QueryString("ID")
sql="Select * from 門派 where 門派='"&Id&"'"
rs=conn.execute(sql)
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: 新細明體}
.c {  font-family: 新細明體; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
</head>
<body bgcolor="#FAF0E2" text="#000000" link="#000080" alink="#800000" vlink="#000080" background="0064.jpg">
<div align="center"> <b><font color="#FF0000" size="+6">門派內容修改</font></b></div>
<form action="updatmp.asp?subject=<%=rs("門派")%>" method=POST>
  <ul>
    <table border=1 cellspacing=0 cellpadding=2 align="center" width="533" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFFF">
      <tr> 
        <td width="99"><font size="2" class="c" color="#FF6600">門派</font></td>
        <td width="420"> 
          <input name="subject" size=40 maxlength=50 value="<%=RS("門派")%>">
          <font color="#FF6600">門派名不得大於10個！ </font></td>
      </tr>
      <tr> 
        <td width="99"><font color="#FF6600">簡介</font></td>
        <td width="420"> 
          <input name="comment" value="<%=RS("簡介")%>" size=40 maxlength=50>
          <font color="#FF6600">字數不超過50個！</font></td>
      </tr>
      <tr> 
        <td width="99"><font class="c" size="2" color="#FF6600">掌門&nbsp;</font></td>
        <td width="420"> 
          <input name="name" value="<%=RS("掌門")%>" size=40 maxlength=50>
          <font color="#FF6600">不設掌門請填寫未定！</font> </td>
      </tr>
      <tr> 
        <td width="99"><font class="c" size="2" color="#FF6600">幫派基金</font></td>
        <td width="420"> 
          <input name="jijing" value="<%=RS("幫派基金")%>" size=40 maxlength=50>
          <font color="#FF6600"> 不得超20億！</font> </td>
      </tr>
      <tr> 
        <td width="99"><font color="#FF6600">適合</font></td>
        <td width="420"> 
          <input name="shihe" value="<%=RS("適合")%>" size=40 maxlength=100>
          <font color="#FF6600">(男/女/所有)</font> </td>
      </tr>
    </table>
    <div align="center"> 
      <p><font size="3" class="c" color="#000000"> <br>
        </font></p>
      <p><font size="3" class="c" color="#000000">一個人隻能當一個門派的掌門，不能為多家的啊，以最新的設置為準。<br>
        <input type="HIDDEN" name="action" value="RegSubmit">
        <input type="SUBMIT" name="Submit" value="更新">
        </font></p>
    </div>
  </ul>
</form>
</body>
</html>
<%
set rs=nothing
%> 
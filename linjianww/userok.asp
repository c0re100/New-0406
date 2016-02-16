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
if info(2)<10   then Response.Redirect "../error.asp?id=900"%>
<html>
<head>
<title>賭局管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: 新細明體}
.c {  font-family: 新細明體; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
</head>
<body bgcolor="#FFFFFF" class=p150 background="0064.jpg">
<div align="center">
  <h1><font color="#FF0000" size="+6">【賭局管理】</font></h1>
</div>
<table width="390" border="1" cellspacing="0" bordercolorlight="#999999" bordercolordark="#FFFFFF" cellpadding="3" align="center">
<tr bgcolor="#0099FF">
    <td width="388"><font face="Wingdings" color="#FFFFFF">1</font><font color="#FFFFFF">賭局清理（用於非正常不能賭博時）</font></td>
</tr>
<tr>
    <td class=p150 width="388" height="34"> ◇ <a href="userok1.asp">銀兩賭局清理</a><br>
      <br>
      ◇ <a href="userok2.asp">法力賭局清理</a><br>
      <br>
      ◇ <a href="userok3.asp">點數賭局清理</a><br>
<br>
      ◇ <a href="userok4.asp">金幣賭局清理</a>  </td>
</tr>
</table>
<br>
</body>
</html> 
<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)<10   then Response.Redirect "../error.asp?id=900"
%>
<html>
<head><style type=text/css>
<!--
body,table {font-size: 9pt; font-family: 新細明體}
.c {  font-family: 新細明體; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
<title>ip管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head>

<body text="#000000" background="0064.jpg" vlink="#0000FF" alink="#0000FF">
<div align="center">
  <p><font size="+6" color="#FF0000">IP管理系統</font></p>
<p>如果在聊天室裡有人常常搗亂可以進行IP管理! <br>
<br>
<a href="ipseek.asp">Ip查找程序</a><br>
<br>
<p><a href="manlock.asp">封IP地址30分鐘</a></p>
<br>
要想永久封ip請改動文件：config.asp！
</div>
</body>
</html>

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
body,table {font-size: 9pt; font-family: �s�ө���}
.c {  font-family: �s�ө���; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
<title>ip�޲z</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head>

<body text="#000000" background="0064.jpg" vlink="#0000FF" alink="#0000FF">
<div align="center">
  <p><font size="+6" color="#FF0000">IP�޲z�t��</font></p>
<p>�p�G�b��ѫǸ̦��H�`�`�o�åi�H�i��IP�޲z! <br>
<br>
<a href="ipseek.asp">Ip�d��{��</a><br>
<br>
<p><a href="manlock.asp">��IP�a�}30����</a></p>
<br>
�n�Q�ä[��ip�Ч�ʤ��Gconfig.asp�I
</div>
</body>
</html>

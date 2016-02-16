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
<title>控制中心</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head>
<body bgcolor="#FFFFFF" background="0064.jpg">
<p><font size="5" face="楷體_GB2312" color="#3333FF">歡迎您使用靈劍江湖管理系統！</font></p>
<p>在此，特感謝所有支持我們江湖發展的朋友，及提出意見的網友！</p>
<p>任何人在沒有作者的書面同意下請不要更改程序和版權！</p>
<p>如有疑問可向該程序設計人員<a href="mailto:Seven.s-888@yahoo.com.tw">江湖總站</a>電郵:Seven.s-888@yahoo.com.tw
------------2003年2月18日--------------</p>
</body>
</html>

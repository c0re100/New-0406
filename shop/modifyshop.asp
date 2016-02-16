<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<head>
<title>修改物品</title>
</head>

<body BACKGROUND="../linjianww/0064.jpg">
<p align="center"><font color="#800080" size="6">調整物品價格</font></p>
<form method="POST" action="man.asp"> 
<div align="center"> <center> <table border="1" width="44%" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolordark="#FFFFFF"> 
<tr> <td width="39%">商店名:</td><td width="66%"><input type="text" name="shopname" style="border: 1px solid; background-color:#EEFFEE;font-size: 9pt; border-color:
#000000 solid" size="10" ></td></tr> <tr> <td width="105%" colspan="2"> <p align="center"><input type="submit" value="提          交" name="B1"></td></tr> 
</table></center></div></form><DIV ALIGN="CENTER"><A HREF="index.asp">返回</A></DIV>
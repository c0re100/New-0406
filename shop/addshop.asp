<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(0)<>"江湖總站"  then  response.redirect "../error.asp?id=425"
%>
<head>
<title>添加商店</title>
</head>

<body bgcolor="#99CCFF">
<p align="center"><font color="#800080" size="6">添加商店</font></p><form method="POST" action="addshop1.asp"> 
<div align="center"> <center> <table border="1" width="70%" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolordark="#FFFFFF"> 
<tr> <td width="39%">商店名:</td><td width="66%"><input type="text" name="shopname" size="10" style="background-color: #99CCFF; color: #000000; border-style: solid; border-width: 1"></td></tr> 
<tr> <td width="39%">店&nbsp; 主:</td><td width="66%"><input type="text" name="name" size="10" style="background-color: #99CCFF; color: #000000; border-style: solid; border-width: 1"></td></tr> 
<tr> <td width="39%">說&nbsp; 明:</td><td width="66%"><input type="text" name="memo" size="33" style="background-color: #99CCFF; color: #000000; border-style: solid; border-width: 1"></td></tr> 
<tr> <td width="39%">經營項目:</td><td width="66%"><select size="1" name="shoptype"> 
<option value="請選擇" selected>請選擇</option> <option value="物品">物品</option> </select></td></tr> 
<tr> <td width="39%">資金:</td><td width="66%"><input type="text" name="shopmoney" size="10" style="background-color: #99CCFF; color: #000000; border-style: solid; border-width: 1"></td></tr> 
<tr> <td width="105%" colspan="2"> <p align="center"><input type="submit" value="提          交" name="B1"></td></tr> 
</table></center></div></form>

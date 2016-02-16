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
<!--#include file="homecheck.asp"-->
<html>

<head>
<style>input, body, select, td { font-size: 14; line-height: 160% }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>體息</title></head>

<body bgcolor="#000000" background="../ljimage/11.jpg">
<p align="center" style="font-size:16;color:yellow"><font color="#000000"><%=info(0)%>您可以放心的在家中睡大覺了：）~~
</font>
<table border="1" bgcolor="#FFCC99" align="center" width="350" cellpadding="10"
cellspacing="13">
<tr>
<td bgcolor="#BEE0FC">
<table>
<tr>
<td> <font color="#000000"><form method="POST" action="sleep1.asp">
你想要休息：
<select name="date" size="1">
<option value="0">零天
<option value="1">一天
<option value="2">兩天
<option value="3">三天
</select>
<select name="time" size="1">
<option value="0">0小時
<option value="1">1小時
<option value="2">2小時
<option value="3">3小時
<option value="4">4小時
<option value="5">5小時
<option value="6">6小時
<option value="7">7小時
<option value="8">8小時
<option value="9">9小時
<option value="10">10小時
<option value="11">11小時
<option value="12">12小時
<option value="13">13小時
<option value="14">14小時
<option value="15">15小時
<option value="16">16小時
<option value="17">17小時
<option value="18">18小時
<option value="19">19小時
<option value="20">20小時
<option value="21">21小時
<option value="22">22小時
<option value="23">23小時
</select>
</font></td>
</tr>
<tr>
<td colspan="2" align="center"><font color="#000000">
<input type="submit" value="確定">
<input
type="reset" value="取消">
</font></td>
</tr>
<tr>
<td colspan="2" style="font-size:9pt">
<hr>
<font color="#000000"> 1、在家中休息可以保護帳號，並可以增加內力和生命值；<br>
2、不收費！按房屋等級不同增加的值不同（詳情見說明）；<br>
3、請準確計算你下次使用帳號的時間！</font></td>
</tr>
<font color="#000000"></form></font>
</table>
</table>

</body>

</html>

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
<title>帳號管理</title>
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
  <h1><font color="#FF0000" size="+6">【帳號管理】</font></h1>
</div>
<table width="390" border="1" cellspacing="0" bordercolorlight="#999999" bordercolordark="#FFFFFF" cellpadding="3" align="center">
<tr bgcolor="#0099FF">
<td width="388"><font face="Wingdings" color="#FFFFFF">1</font><font color="#FFFFFF">清理帳號</font></td>
</tr>
<tr>
    <td class=p150 width="388" height="34"> ◇ <a href="manaccdel7.asp?page=1">清理７天內隻用過一次的帳號</a><br>
◇ <a href="maingl.asp">更新所有用戶管理等級為1(掌門5級管理員不變)</a><br>
◇ <a href="maingl1.asp">更新所有的人配偶為無</a><br>
◇ <a href="maingl2.asp">更新所有現金/銀兩為負用戶為0</a><br>
◇ <a href="maingl3.asp">發會費（非官府）</a><br>
◇ <a href="maingl4.asp">發銀兩1億（官府）</a><br>
◇ <a href="maingl5.asp">月積分前20名獎勵50會費</a><br>
◇ <a href="mmm1.asp">更新所有監禁狀態的人為正常</a>
</td>
</tr>
</table>
<br>
<br>
<br>
<table width="390" border="1" cellspacing="0" bordercolorlight="#999999" bordercolordark="#FFFFFF" cellpadding="3" align="center" height="75">
<tr bgcolor="#0099FF">
<td width="388"><font face="Wingdings" color="#FFFFFF">1</font><font color="#FFFFFF">帳號查詢</font></td>
</tr>
<tr>
    <td class=p150 width="388" height="44"> ◇ <a href="manaccqueryvalue.asp?page=1">按累積分列出帳號</a><br>
◇ <a href="manaccquerymvalue.asp?page=1">按月積分列出帳號</a><br>
◇ <a href="manaccquerygrade.asp">按級別列出所有帳號</a></td>
</tr>
</table>
<br>
<hr noshade size="1" color=009900>
<b> 注意事項 </b><br>
帳號即ID，是用戶進入江湖時注冊的昵稱。本功能用於清理３０天未用的帳號、７天來隻用過一次的帳號、已經自殺的用戶名；設定/解除永久保留某些用戶名；更改某個用戶名的密碼（用於用戶遺忘密碼時，站長將其密碼重新設定，然後通知該用戶新密碼）；查詢某個用戶名的全部資料；按級別列出所有用戶名。
</body>
</html> 
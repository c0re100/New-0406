<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)<6  then response.redirect "../error.asp?id=425"
%>
<html>
<head>
<title>出題答對有獎(感謝喬峰)</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<script language=javascript>window.moveTo(100,50);window.resizeTo(screen.availWidth*2/3,screen.availHeight*3/4);</script>
</head>
<%
%>
<form method="post" action="asking.asp">
<table width="250" border="1" cellspacing="0" cellpadding="3" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr align="center" bgcolor="#33CCCC">
<td colspan="2" height="27">搶答系統出題</td>
</tr>
<tr>
<td width="32" valign="top">問：</td>
<td width="202">
<textarea name="ask" rows="3" wrap="VIRTUAL" cols="30"></textarea>
</td>
</tr>
<tr>
<td width="32">答：</td>
<td width="202">
<input type="text" name="reply" size="15" maxlength="15">
</td>
</tr>
<tr>
<td> 獎：</td>
<td>
<input type=text name='silver' value='1000000' maxlength=7 size=7>
</td>
</tr>
<tr>
<td colspan="2">
<div align="center">
<input type="submit" name="Submit" value="提 交">
<input type="submit" name="Submit2" value="關 閉">
</div>
</td>
</tr>
<tr>
<td colspan="2">
<div align="center">出題人：<%=info(0)%> </div>
</td>
</tr>
</table>
</form>
<p align="center">網管負責出題，對於答對的聊友，系統將給予獎勵！<br>
注意!不得利用此功能做假! </p>
</body>
</html>

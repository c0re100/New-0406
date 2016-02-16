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
%><html>
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
<font color="#FF0000">【按級別列出所有帳號】</font></div>
<div align="center"><a href="txt.asp">返回</a> </div>
<hr noshade size="1" color=009900>
<b> 注意事項 </b><br>
請選擇要查詢的級別。
<hr noshade size="1" color=009900>
<div align=center>
<table border="1" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center" bgcolor="#E0E0E0" cellpadding="8">
<form method="post" action="manaccquerygradeok.asp">
<tr align="center">
<td>
<table width="100%" border="0" cellpadding="4">
<tr>
<td>要查詢的級別：</td>
<td>
<select name="grade" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
<option value="0">0 被禁用</option>
<option value="1" selected>1 級</option>
<option value="2">2 級</option>
<option value="3">3 級</option>
<option value="4">4 級</option>
<option value="5">5 級</option>
<option value="6">6 級</option>
<option value="7">7 級</option>
<option value="8">8 級</option>
<option value="9">9 級</option>
<option value="10">10 級</option>
</select>
</td>
</tr>
<tr>
<td colspan="2" align="center">
<input type="submit" name="Submit" value="查詢">
<input type="reset" value="重寫" name="reset">
<input type="button" value="放棄" onClick="javascript:history.go(-1)" name="button">
</td>
</tr>
</table>
</td>
</tr>
</form>
</table>
</div>
</body>
</html>
 
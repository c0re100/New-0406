<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Session("iq")="" then
response.redirect "index.asp"
end if
%>
<HTML>
<HEAD>
<title>第四關：賭大小</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">

<LINK href="../../chat/READONLY/Style.css" rel=stylesheet></HEAD>

<body bgcolor="#000000" oncontextmenu=self.event.returnValue=false background="images/bg.gif" >
<div align="left"></div>

<div align=center>
<table width="778" border="0" cellspacing="0" cellpadding="0">
<tr>
<td><font size="2" class=c><font size="3"><b> </b></font></font>
<div align=center><b>司徒輸光</b> <font color="#FF0033">vs</font> <b><%=ljjh_nickname%></b></div>

</td>
</tr>
</table>
<table width="778" border="0" cellspacing="0" cellpadding="0">
<tr>
<td>
<div align="center">一局定輸贏呵，考慮清楚。</div>
</td>
</tr>
</table>
<form method="POST" action="b&amp;spose.asp">
<table border=1 cellspacing=0 cellpadding=0 align="center" width="350" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr align="center" bgcolor="#FFFFFF">
<td width="33%"><font size="2" color="#000000"><img src="run.gif" width="38" height="36"></font></td>
<td width="33%"><font size="2" color="#000000"><img src="run.gif" width="38" height="36"></font></td>
<td width="33%"><font size="2" color="#000000"><img src="run.gif" width="38" height="36"></font></td>
</tr>
<tr bgcolor="#336633">
<td width="960" colspan="3" height="29"><font size="2" class="c" color="#000000"><b>&nbsp;&nbsp;<font color="#FFFFFF">請選擇</font></b></font></td>
</tr>
<tr bgcolor="#FFFFFF">
<td align=center colspan="3">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
<tr align="center">
<td width="50%"><img src="big.gif" width="46" height="40"></td>
<td width="50%"><img src="small.gif" width="46" height="40"></td>
</tr>
<tr align="center">
<td width="50%">
<input type="radio" name="select" value="big" checked>
</td>
<td width="50%">
<input type="radio" name="select" value="small">
</td>
</tr>
</table>
</td>
</tr>
<tr>
<td align=center colspan="3">&nbsp;</td>
</tr>
<tr>
<td align=center colspan="3"> <font size="2" color="#000000">
<input type="submit" value="來吧，看誰好運。" name="B1" >
<input type="reset" value="我要考慮一下：）" name="B2" >
</font></td>
</tr>
</table>
</form>
</div>
</BODY>
</HTML>
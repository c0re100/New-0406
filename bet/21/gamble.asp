<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
%>
<!--#include file="h_gamble.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
dim msg
	
select case Session("n_Result")
case 1
msg="你贏了"
case 2
msg="你輸了"
case 3
msg="平手呀"
case 4
msg=""
case else
msg=""
end select
%>
<html>

<head>
<title>賭場</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body bgcolor="#FFFFFF" text="#000000" topmargin="0">
<table width="95%" border="0" cellspacing="0" cellpadding="0" bgcolor="#669900"
align="center" height="277">
<tr>
<td width="22" height="21" align="left" valign="top"><img
src="jhimg/du1.gif" width="22" height="21"></td>
<td rowspan="3" valign="top">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="68%" height="28">
<div align="center"> <img src="jhimg/lswjs.gif" width="244" height="52">
</div>
</td>
<td width="32%" valign="bottom" height="28"><img
src="jhimg/duren.gif" width="75" height="54"></td>
</tr>
<tr align="left" valign="top">
<td colspan="2" height="50">
<table width="90%" border="0" cellspacing="1" cellpadding="0"
align="center">
<tr>
<td width="71" height="61"><%=GejhimgURL(Session("n_PokerBack"))%></td>
<td height="61"></td>
</tr>
</table>
<table width="90%" border="0" cellspacing="1" cellpadding="0"
align="center">
<tr>
<td width="71" height="69">&nbsp;</td>
<td height="69" align="left">
<%call ShowAntiPokers()%>
</td>
</tr>
<tr>
<td>&nbsp;</td>
<td>
<%call ShowUserPokers()%>
</td>
</tr>
</table>
</td>
</tr>
<tr>
<td colspan="2"></td>
</tr>
<tr>
<td colspan="2" height="2" bgcolor="#7bc621">
<form method="post" action="do_gamble.asp">
<table width="100%" border="0" cellspacing="1" cellpadding="0">
<tr>
<td colspan="1">當前點數：<%=GetMaxPoint(Session("n_UserPokers"))%></td>
<td colspan="2">資金總數：<%=CCur(Session("n_TotalMoney"))%></td>
<td colspan="2">&nbsp;</td>
</tr>
<tr>
<td>我押：<%=Session("n_Bet")%></td>
<td> <%
if Session("n_Begin") then
Response.Write "<b>" & Session("n_Bet") & "</b>"
else
%>
<input type="text" maxlength="10" name="Bet" value="0"
size="10"
<%
if ccur(session("n_totalmoney"))=0 then response.write enablebtn(false)
%>>
<%
end if
%> <%=msg%> </td>
<td>
<%
if not Session("n_Begin") then
if Session("n_Init") then
if CCur(Session("n_TotalMoney"))<>0 then
%>
<input type="submit" name="Action" value="再來一局">
								
<%
else
%>
<input type="submit" name="Action" value="返回">
<%
end if
else
%>
<input type="submit" name="Action" value="開局">
<%
end if
end if
%>
</td>
<td><input type="submit" name="Action" value="要牌"
<%=enablebtn(session("n_begin"))%>></td>
<td><input type="submit" name="Action" value="開牌"
<%=enablebtn(session("n_begin"))%>></td>
</tr>
</table>
</form>
</td>
</tr>
<tr>
<td colspan="2" height="10">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td height="2">
<div align="center"><a href="../../jh.asp">返回</a> </div>
</td>
</tr>
</table>
</td>
</tr>
<tr>
<td colspan="2"></td>
</tr>
</table>
</td>
<td width="27" height="21" align="right" valign="top"><img
src="jhimg/du2.gif" width="27" height="21"></td>
</tr>
<tr>
<td height="208" width="4%" bgcolor="#669900"> </td>
<td height="208" bgcolor="#669900" width="4%"> </td>
</tr>
<tr>
<td width="22" height="41" align="left" valign="bottom"><img
src="jhimg/du3.gif" width="32" height="23"></td>
<td width="27" height="41" align="right" valign="bottom"><img
src="jhimg/du4.gif" width="27" height="23"></td>
</tr>
</table>
</body>

</html>
<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")

if Application("ljjh_ask")="" then
response.redirect "../error.asp?id=476"
response.end
end if
username=info(0)
mycorp=info(5)
if mycorp="�x��" then response.redirect "../error.asp?id=477"
%>
<html>
<head>
<title>���靈��(�P�³�p)</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<script language=javascript>window.moveTo(100,50);window.resizeTo(screen.availWidth*2/3,screen.availHeight*3/4);</script>
</head>
<%
%>
<form method="post" action="replying.asp">
<table width="285" border="1" cellspacing="0" cellpadding="3" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr align="center" bgcolor="#33CCCC">
<td colspan="2" height="27">�m����</td>
</tr>
<tr>
<td width="53" valign="top">�ݡG</td>
<td width="214">
<textarea name="ask" readonly rows="3" wrap="VIRTUAL" cols="30"><%=Application("ljjh_ask")%></textarea>
</td>
</tr>
<tr>
<td width="53">���G</td>
<td width="214">
<input type="text" name="reply" size="10" maxlength="225">
</td>
</tr>
<tr>
<td width="53">���y�G</td>
<td width="214"> <%=Application("ljjh_asksilver")%>��</td>
</tr>
<tr>
<td colspan="2">
<div align="center">
<input type="submit" name="Submit" value="�� ��">
<input type="submit" name="Submit2" value="�� ��">
</div>
</td>
</tr>
<tr>
<td colspan="2">
<div align="center">���D�H�G<%=info(0)%> </div>
</td>
</tr>
</table>
</form>
</body>
</html>

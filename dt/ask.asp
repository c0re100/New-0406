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
<title>�X�D���靈��(�P�³�p)</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<script language=javascript>window.moveTo(100,50);window.resizeTo(screen.availWidth*2/3,screen.availHeight*3/4);</script>
</head>
<%
%>
<form method="post" action="asking.asp">
<table width="250" border="1" cellspacing="0" cellpadding="3" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr align="center" bgcolor="#33CCCC">
<td colspan="2" height="27">�m���t�ΥX�D</td>
</tr>
<tr>
<td width="32" valign="top">�ݡG</td>
<td width="202">
<textarea name="ask" rows="3" wrap="VIRTUAL" cols="30"></textarea>
</td>
</tr>
<tr>
<td width="32">���G</td>
<td width="202">
<input type="text" name="reply" size="15" maxlength="15">
</td>
</tr>
<tr>
<td> ���G</td>
<td>
<input type=text name='silver' value='1000000' maxlength=7 size=7>
</td>
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
<div align="center">�X�D�H�G<%=info(0)%> </div>
</td>
</tr>
</table>
</form>
<p align="center">���ޭt�d�X�D�A��󵪹諸��͡A�t�αN�������y�I<br>
�`�N!���o�Q�Φ��\�వ��! </p>
</body>
</html>

<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Session("iq")<>"asw" then
response.redirect "index.asp"
end if
%>
<html>
<head>
<title>�ĤG��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<LINK href="../style.css" rel=stylesheet></head>

<body bgcolor="#000000" oncontextmenu=self.event.returnValue=false background="images/bg.gif" >
<table width="600" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td width="595">
      <div align="left"><b><font color="#0000FF">�ĤG���O�@��PPMM�C~�I�I�I</font></b></div>
</td>
<td width="165" rowspan="2" valign="top"><font color="#FFFFFF"><img src="jh0.jpg" width="153" height="172"></font></td>
</tr>
<tr>
<td width="595" valign="top">
      <p style="line-height: 200%; margin: 20"> PPMM�G�١A�p�ӭ��A�ڤ��ˬO�W�t�������n�C�ڨӥN�ڪ����˧��ĤG���A���n�ȧr�A�ڳo�̦��Q�D���D�A��������F�N�i�H�L�h�A�n���n�ոէr�H</p>
</td>
</tr>
<tr valign="top">
<td colspan="2">
<p><font color=red><%=info(0)%>:</font><a href="answer.asp">�n�a�A�ڪ��~���i�O�s�������D�Ѥ~���A����l���ڡC</a><br>
</p>
</td>
</tr>
</table>
</body>
</html>
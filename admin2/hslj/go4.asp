<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Session("iq")<>"opy" then
response.redirect "index.asp"
end if %>
<html>
<head>
<title>�ĤT��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<LINK href="../style.css" rel=stylesheet></head>

<body oncontextmenu=self.event.returnValue=false background="images/bg.gif" >
<table width="600" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td width="599">
      <div align="left"><font color="#FFFFFF"><b><font color="#0000FF">�A����ĤT�����f�e�A��M���X�@�ӥ��ꪺ�_�ΩǪ��p���Y~</font></b></font></div>
</td>
<td width="147" rowspan="2" valign="top"><font color="#FFFFFF"><img src="face.gif" width="142" height="241"></font></td>
</tr>
<tr>
<td width="599" valign="top">
      <p style="line-height: 200%; margin: 20"> �������G��p�l�A���Ǥ������C�w�g��F�ĤT���C�ڥs�������A�q�T���}�l���[�A�줵�Ѥw�g60�h�~�F�C�Q�L�ڳo���A�N�ʰʧA�����Y�A���ڦ��Y���������C�n�O�A�������̰��L�ڡA�ڴN��A�L�h~�I
      <p><font color="#FFFFFF"><%=info(0)%>:<a href="go5.asp">�ȧA����A�ڪZ�\�i�O�ѤU�Ĥ@�I�I</a>
</font>
</p>
<p>
</p>
<p> </p>
</td>
</tr>
</table>
</body>
</html>

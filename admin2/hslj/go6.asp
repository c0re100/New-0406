<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if Session("iq")="" then
response.redirect "index.asp"
end if
%>
<html>
<head>
<title>�ĥ|��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">

<LINK href="../../chat/READONLY/Style.css" rel=stylesheet></head>

<body oncontextmenu=self.event.returnValue=false background="images/bg.gif" >
<table width="600" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td width="617">
      <div align="left"><b><font color="#0000FF">�訫��ĥ|�����e�A�@�ӿW���s�N�q��᭱��F�X��  </font></b></div>
</td>
<td width="136" rowspan="2" valign="top"><font color="#FFFFFF"><img src="face2.GIF" width="136" height="283"></font></td>
</tr>
<tr>
<td width="617" valign="top">
      <p style="line-height: 200%; margin: 20"> �q�{����G�̤p�l�A�Q���L�ڥq�{����o�@���i�S����e���C�ӨӨӡA���̨ӽ��@��A�n�O�A��F�A������Ӫ��N�^����h�F�n�O�AĹ�F�A�ڴN��A�L�h�C���ˡA�n���n�ոաH 
      <p style="line-height: 200%; margin: 20"><font color="#FFFFFF"><%=info(0)%>:<a href="go7.asp">�Ѥl�i�O�ѤU�Ĥ@�䯫�A�ȧA����C</a><br>
</font>
</p>

</td>
</tr>
</table>
</body>
</html>
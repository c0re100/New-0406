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
<html>
<head>
<title>�̫�@��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<LINK href="../chat/READONLY/Style.css" rel=stylesheet></head>

<body bgcolor="#000000" oncontextmenu=self.event.returnValue=false background="images/bg.gif" >
<table width="600" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td width="630"> 
      <div align="left"><b><font color="#000000">�̫�@���F�A�_�ê����e���ۤ@�Ӧ~���H  </font></b></div>
    </td>
    <td width="128" rowspan="2" valign="top"><img src="face3.GIF" width="128" height="290"></td>
  </tr>
  <tr> 
    <td width="630" valign="top"> 
      <p style="line-height: 200%; margin: 20"> <font color="#9966CC">�~���H�G�O�ݧڬO��~�I�A���n���D�ڬO�A���̫�@�ӹ��N�i�H�F�C�e����4���A�A�g�L�F�ѤO�B���O�B�����B��N������A�{�b���ڨӦ���A�����O�A�A�������ѡA�~��Τ��O�}�}�_�ê��j���A�񰨹L�ӧa�A�ڤ��|�d�����I</font> 
      <p style="line-height: 200%; margin: 20"><font color="#FFFFFF"><%=info(0)%>:<a href="go9.asp">�n~�ǳƦn�鵹�ڧa�I�I�I</a><br>
        </font> </p>
      <p style="line-height: 200%; margin: 20"> 
    </td>
  </tr>
  <tr valign="top"> 
    <td colspan="2"> </td>
  </tr>
</table>
</body>
</html>
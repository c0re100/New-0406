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
if info(2)<10   then Response.Redirect "../error.asp?id=900"%>
<html>
<head>
<title>�b���޲z</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: �s�ө���}
.c {  font-family: �s�ө���; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
</head>
<body bgcolor="#FFFFFF" class=p150 background="0064.jpg">
<div align="center">
  <h1><font color="#FF0000" size="+6">�i�b���޲z�j</font></h1>
</div>
<table width="390" border="1" cellspacing="0" bordercolorlight="#999999" bordercolordark="#FFFFFF" cellpadding="3" align="center">
<tr bgcolor="#0099FF">
<td width="388"><font face="Wingdings" color="#FFFFFF">1</font><font color="#FFFFFF">�M�z�b��</font></td>
</tr>
<tr>
    <td class=p150 width="388" height="34"> �� <a href="manaccdel7.asp?page=1">�M�z���Ѥ����ιL�@�����b��</a><br>
�� <a href="maingl.asp">��s�Ҧ��Τ�޲z���Ŭ�1(�x��5�ź޲z������)</a><br>
�� <a href="maingl1.asp">��s�Ҧ����H�t�����L</a><br>
�� <a href="maingl2.asp">��s�Ҧ��{��/�Ȩ⬰�t�Τᬰ0</a><br>
�� <a href="maingl3.asp">�o�|�O�]�D�x���^</a><br>
�� <a href="maingl4.asp">�o�Ȩ�1���]�x���^</a><br>
�� <a href="maingl5.asp">��n���e20�W���y50�|�O</a><br>
�� <a href="mmm1.asp">��s�Ҧ��ʸT���A���H�����`</a>
</td>
</tr>
</table>
<br>
<br>
<br>
<table width="390" border="1" cellspacing="0" bordercolorlight="#999999" bordercolordark="#FFFFFF" cellpadding="3" align="center" height="75">
<tr bgcolor="#0099FF">
<td width="388"><font face="Wingdings" color="#FFFFFF">1</font><font color="#FFFFFF">�b���d��</font></td>
</tr>
<tr>
    <td class=p150 width="388" height="44"> �� <a href="manaccqueryvalue.asp?page=1">���ֿn���C�X�b��</a><br>
�� <a href="manaccquerymvalue.asp?page=1">����n���C�X�b��</a><br>
�� <a href="manaccquerygrade.asp">���ŧO�C�X�Ҧ��b��</a></td>
</tr>
</table>
<br>
<hr noshade size="1" color=009900>
<b> �`�N�ƶ� </b><br>
�b���YID�A�O�Τ�i�J����ɪ`�U����١C���\��Ω�M�z�����ѥ��Ϊ��b���B���ѨӰ��ιL�@�����b���B�w�g�۱����Τ�W�F�]�w/�Ѱ��ä[�O�d�Y�ǥΤ�W�F���Y�ӥΤ�W���K�X�]�Ω�Τ��ѱK�X�ɡA�����N��K�X���s�]�w�A�M��q���ӥΤ�s�K�X�^�F�d�߬Y�ӥΤ�W��������ơF���ŧO�C�X�Ҧ��Τ�W�C
</body>
</html> 
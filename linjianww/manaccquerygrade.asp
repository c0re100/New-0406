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
<font color="#FF0000">�i���ŧO�C�X�Ҧ��b���j</font></div>
<div align="center"><a href="txt.asp">��^</a> </div>
<hr noshade size="1" color=009900>
<b> �`�N�ƶ� </b><br>
�п�ܭn�d�ߪ��ŧO�C
<hr noshade size="1" color=009900>
<div align=center>
<table border="1" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center" bgcolor="#E0E0E0" cellpadding="8">
<form method="post" action="manaccquerygradeok.asp">
<tr align="center">
<td>
<table width="100%" border="0" cellpadding="4">
<tr>
<td>�n�d�ߪ��ŧO�G</td>
<td>
<select name="grade" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
<option value="0">0 �Q�T��</option>
<option value="1" selected>1 ��</option>
<option value="2">2 ��</option>
<option value="3">3 ��</option>
<option value="4">4 ��</option>
<option value="5">5 ��</option>
<option value="6">6 ��</option>
<option value="7">7 ��</option>
<option value="8">8 ��</option>
<option value="9">9 ��</option>
<option value="10">10 ��</option>
</select>
</td>
</tr>
<tr>
<td colspan="2" align="center">
<input type="submit" name="Submit" value="�d��">
<input type="reset" value="���g" name="reset">
<input type="button" value="���" onClick="javascript:history.go(-1)" name="button">
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
 
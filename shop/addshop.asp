<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(0)<>"�����`��"  then  response.redirect "../error.asp?id=425"
%>
<head>
<title>�K�[�ө�</title>
</head>

<body bgcolor="#99CCFF">
<p align="center"><font color="#800080" size="6">�K�[�ө�</font></p><form method="POST" action="addshop1.asp"> 
<div align="center"> <center> <table border="1" width="70%" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolordark="#FFFFFF"> 
<tr> <td width="39%">�ө��W:</td><td width="66%"><input type="text" name="shopname" size="10" style="background-color: #99CCFF; color: #000000; border-style: solid; border-width: 1"></td></tr> 
<tr> <td width="39%">��&nbsp; �D:</td><td width="66%"><input type="text" name="name" size="10" style="background-color: #99CCFF; color: #000000; border-style: solid; border-width: 1"></td></tr> 
<tr> <td width="39%">��&nbsp; ��:</td><td width="66%"><input type="text" name="memo" size="33" style="background-color: #99CCFF; color: #000000; border-style: solid; border-width: 1"></td></tr> 
<tr> <td width="39%">�g�綵��:</td><td width="66%"><select size="1" name="shoptype"> 
<option value="�п��" selected>�п��</option> <option value="���~">���~</option> </select></td></tr> 
<tr> <td width="39%">���:</td><td width="66%"><input type="text" name="shopmoney" size="10" style="background-color: #99CCFF; color: #000000; border-style: solid; border-width: 1"></td></tr> 
<tr> <td width="105%" colspan="2"> <p align="center"><input type="submit" value="��          ��" name="B1"></td></tr> 
</table></center></div></form>

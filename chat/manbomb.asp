<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
grade=Int(info(2))
mypai=info(5)
inthechat=Session("ljjh_inthechat")
if grade<10 then Response.Redirect "../error.asp?id=482"
if inthechat<>"1" then Response.Redirect "manerr.asp?id=482"
bombname=Trim(Request.QueryString("id"))
if bombname="" then Response.Redirect "../error.asp?id=481"
if CStr(bombname)=CStr(info(0)) then Response.Redirect "../error.asp?id=481"%><html>
<head>
<title>���u�ާ@</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="readonly/style.css">
<script language=javascript>window.moveTo(100,50);window.resizeTo(screen.availWidth*2/3,screen.availHeight*3/4);</script>
</head>
<body bgcolor="#FFFFFF" class=p150>
<div align="center">
<h1><font color="0099FF"><font size="6" face="����" color="#000000">�� �u �� �@</font></font></h1>
</div>
<hr noshade size="1" color=009900 width="70%">
<div align="center">
<table width="380" border="0" cellspacing="2" cellpadding="2">
<tr>
<td><b> �`�N�ƶ� </b> <br>
������J�����u����]�~���񬵼u�A�Q�����ﹳ�������N�����a�u�X�s���f�A����t�θ귽�ӺɡB�����C<br>
���u�ާ@�|�Q�O���b����Ȥ��}���椤�A�Ѽs�j��ͺʷ��A�]���A�Ф��H�N���H�I </td>
</tr>
</table>

</div>
<hr noshade size="1" color=009900 width="70%">
<br>
<table border="1" cellspacing="0" cellpadding="3" bgcolor="E0E0E0" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center" width="380">
<form method="post" action="manbombok.asp">
<tr>
<td>
<table border="0" cellpadding="4" width="100%" cellspacing="0">
<tr>
<td align="right" width="26%">�F���Τ�W�G</td>
<td width="74%"><font color="#FF0000"><%=bombname%>
<input type="hidden" name="bombname" value="<%=bombname%>">
</font></td>
</tr>
<tr>
<td align="right" width="26%">�F������]�G</td>
<td width="74%">(&lt;=60�r��) </td>
</tr>
<tr>
<td align="right" width="26%">
<select name="select" onchange="javascript:document.forms[0].bombwhy.value=this.value;document.forms[0].bombwhy.focus();">
<option value="" selected>�۶�</option>
<option value="�����w��H�C">���w��</option>
<option value="�Ҩ����W�r�Q�������C">����ID</option>
<option value="�è�̡Aĵ�i�S��ť�C">�è��</option>
<option value="�b��ѫǴ��������۲z�D�w�����סC">���D�w</option>
<option value="����u��ѳW�h�A�i��H�������C">�|�H</option>
<option value="�b��ѫǴ����H�ϰ�a�k�ߪk�W�����סC">�H�k</option>
</select>
�G</td>
<td width="74%">
<input type="text" name="bombwhy" maxlength="60" size="40">
</td>
</tr><%if info(0)="�����`��" then%>
<tr>
<td align="right" width="26%">�޲z���ﶵ�G</td>
<td width="74%">
<input type="radio" name="logok" value="1" checked>�O�J��Ȥ��}��
<input type="radio" name="logok" value="0">���O�J��Ȥ��}��</td>
</tr><%end if%>
</table>
<div align="center">
<input type="submit" name="bombok" value="�F��">
<input type="button" value="���" onclick="javascript:history.go(-1)">
</div>
</td>
</tr>
</form>
</table>
<br>
<hr noshade size="1" color=009900 width="70%">
<div align=center class=cp></div>
</body>
</html> 
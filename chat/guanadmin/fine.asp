<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Response.Expires=0
if info(2)<8  then Response.Redirect "../../error.asp?id=425"
%>
<html>
<head>
<title>�b�u��H�S�O�޲z</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="../chat/READONLY/STYLE.CSS">
</head>

<body topmargin="0" bgcolor="#0066FF" text="#00FFFF">
<p align="center">�b�u��H�S�O�޲z</p>
<p align="center">�޲z�����G<br>
  <font color="yellow"><br>
  �{�ھڥH�U���ܿ�J�W�r�B��H��]�I������Y�i�A�p�������A�Ч䦿���`���pô</font></p>
<div align="center"></div>
<form method="POST" action="zhixin.asp">
  <div align="center">��ܺ޲z���<br>
    ��J�n��H���W�r�B��]<br>
    <select name="ljgl">
      <option value="�e��" selected>�e��</option>
      <option value="���c">���c</option>
      <option value="�ʸT">�ʸT</option>
      <option value="�_��">�_��</option>
      <option value="���u">���u</option>
    </select>
    <input type="text" name="pass1" size="10" maxlength="10">
    <input type="text" name="pass11" size="10" maxlength="10">
    <input type="submit" value="����" name="a1" class="p9">
    <input type="reset" name="Submit22" value="����">
    <br>
  </div>
</form>
</body>
</html>
 
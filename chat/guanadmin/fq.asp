<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Response.Expires=0
if info(2)<7  then Response.Redirect "../error.asp?id=439"
%>
<html>
<head>
<title>�b�u�o���޲z�B</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="../READONLY/STYLE.CSS">
</head>

<body topmargin="0" bgcolor="#660000" text="#00FFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" leftmargin="0">
<p align="center">�b�u�o���޲z�B</p>
<div align="center"></div>
<table width="111" border="0" align="center" cellspacing="4" cellpadding="4">
  <tr> 
    <td><a href="sendfy.asp" target="_blank"><font color="#FFFFFF">�o�Ȩ�</font></a></td>
  </tr>
  <tr> 
    <td><a href="sendfn.asp" target="_blank"><font color="#FFFFFF">�o���O</font></a></td>
  </tr>
  <tr> 
    <td><a href="sendft.asp" target="_blank"><font color="#FFFFFF">�o��O</font></a></td>
  </tr>
  <tr> 
    <td><a href="sendff.asp" target="_blank"><font color="#FFFFFF">�o�k�O</font></a></td>
  </tr>
  <tr> 
    <td><a href="sendfj.asp" target="_blank"><font color="#FFFFFF">�o����</font></a></td>
  </tr>
</table>
</body>
</html>
 
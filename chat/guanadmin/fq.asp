<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Response.Expires=0
if info(2)<7  then Response.Redirect "../error.asp?id=439"
%>
<html>
<head>
<title>在線發錢管理處</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="../READONLY/STYLE.CSS">
</head>

<body topmargin="0" bgcolor="#660000" text="#00FFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" leftmargin="0">
<p align="center">在線發錢管理處</p>
<div align="center"></div>
<table width="111" border="0" align="center" cellspacing="4" cellpadding="4">
  <tr> 
    <td><a href="sendfy.asp" target="_blank"><font color="#FFFFFF">發銀兩</font></a></td>
  </tr>
  <tr> 
    <td><a href="sendfn.asp" target="_blank"><font color="#FFFFFF">發內力</font></a></td>
  </tr>
  <tr> 
    <td><a href="sendft.asp" target="_blank"><font color="#FFFFFF">發體力</font></a></td>
  </tr>
  <tr> 
    <td><a href="sendff.asp" target="_blank"><font color="#FFFFFF">發法力</font></a></td>
  </tr>
  <tr> 
    <td><a href="sendfj.asp" target="_blank"><font color="#FFFFFF">發金幣</font></a></td>
  </tr>
</table>
</body>
</html>
 
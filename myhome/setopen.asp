<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<!--#include file="homecheck.asp"-->
<!--#include file="data1.asp"-->
<html>

<head>
<title>�F�C����-�]�I�v��</title>
<link rel="stylesheet" type="text/css" href="../style.css">
<style type="text/css">td           { font-family: �s�ө���; font-size: 9pt }
body         { font-family: �s�ө���; font-size: 9pt }
select       { font-family: �s�ө���; font-size: 9pt }
a            { color: #FFC106; font-family: �s�ө���; font-size: 9pt; text-decoration: none }
a:hover      { color: #cc0033; font-family: �s�ө���; font-size: 9pt; text-decoration:
underline }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body bgcolor="#000000" text="#000000" background="../ljimage/11.jpg" link="#0000FF">
<table width="100%" border="0" cellspacing="0" cellpadding="0"
bordercolorlight="#000000" bordercolordark="#FFFFFF"
align="center">
<tr>
<td width="91" height="26"><font size="2" color="#000000">&nbsp; <span class="zilong">
<%
y=Month(date())
r=Day(date())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
Response.Write  y & "��" & r & "��" %>
</span></font></td>
<td width="475" height="26">
<div align="center"><font color="#000000"> <font size="+2">����]�I�v��</font></font></div>
</td>
<td width="104">
<div align="center"> </div>
</td>
</tr>
</table>
<br>
<table width="738" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td width="17">&nbsp;</td>
<td width="47" valign="top">
<div align="center"> </div>
</td>
<td valign="top" width="639">
<div align="center">
<div align="center">
<div align="center">
<table border="0" width="468" cellspacing="1" cellpadding="0"
height="20">

</table>
</div>
</div>
<div align="center">
<center>
<table border="0" width="537" cellspacing="1" cellpadding="0">
<tr>
<td width="500">
<form method="POST" action="setopen_post.asp">
<table border="1" width="497" cellspacing="1" cellpadding="0"
bordercolor="#0000FF">
<tr>
<td class="p3" width="154"><font color="#000000">�ڪ���O</font></td>
<td class="p3" width="79">
<p><font color="#000000">
<input type="radio" value="�O" checked name="diary">
�}��</font></p>
</td>
<td class="p3" width="264"><font color="#000000">
<input type="radio" value="�_"
name="diary">
����</font></td>
</tr>
<tr>
<td class="p3" width="154"><font color="#000000">�ڪ���ñ</font></td>
<td class="p3" width="79"><font color="#000000">
<input type="radio" value="�O"
checked name="bookmark">
�}��</font></td>
<td class="p3" width="264"><font color="#000000">
<input type="radio" value="�_"
name="bookmark">
����</font></td>
</tr>
<tr>
<td class="p3" width="154"><font color="#000000"> </font></td>
<td class="p3" width="79"><font color="#000000"> </font></td>
<td class="p3" width="264"><font color="#000000">
<input type="submit"
value="����" name="B1">
<input type="reset"
value="���g" name="B2">
</font></td>
</tr>
</table>
</form>
</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="500" cellspacing="1" cellpadding="0">
<tr>
<td class="p1" width="500">&nbsp;&nbsp;</td>
</tr>
</table>
</center>
</div>
</div>
</td>
<td width="25"> </td>
</tr>
</table>
<div align="center">
</div>

</body>

</html>
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
<!--#include file="data2.asp"-->
<html>

<head>
<title>�F�C����-���X�ȯd��</title>
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
<td width="91" height="26"><font size="2">&nbsp; <font
color="#FFFFFF"></font><font size="2"><font color="#ffffff"
size="2"><span class="zilong"><font color="#FFCC00">
<%
y=Month(date())
r=Day(date())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
Response.Write  y & "��" & r & "��" %>
</font></span></font></font></font></td>
<td width="475" height="26">
<div align="center"> <font size="2" color="#E18C59"><b><font color="#000000" size="+2"> ���X�ȯd��</font></b></font>
</div>
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
<td valign="top" width="639">
<div align="center" style="width: 689; height: 41">
<div align="center">
<div align="center">
<table border="0" width="468" cellspacing="1" cellpadding="0"
height="20">

</table>
</div>
</div>
<div align="center">
<center>
<table border="1" width="695" cellspacing="1" cellpadding="0"
bordercolor="#000000">
<tr>
<td class="p1" width="336">��-�z��e����m- --�L�ȴJ��--�ӤH�H��--���X�ȯd��</td>
<td class="p1" width="440">��-��e�ɶ��G<%=date%><%=time%></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<form method="POST" action="setguest_post.asp">
<table border="1" width="690" cellspacing="1" cellpadding="0"
bordercolor="#000000">
<tr>
<td class="p3" width="10"> </td>
<td class="p3" width="326">���ӳX�̪���<br>
<textarea name="guestmemo" rows="7" cols="43"></textarea>
<td class="p3" width="440" valign="bottom"><input type="submit"
value="����" name="B1"><input type="reset"
value="�������g" name="B2"></td>
</tr>
</table>
</form>
</center>
</div>
<div align="center"> <a href="javascript:history.back(1)"><font color="#000000">��^</font></a>
<center>
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
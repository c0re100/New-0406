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

 
<!--#include file="../data1.asp"-->
<html>

<head>
<title>�F�C����-�d���ө�</title>
<link rel="stylesheet" type="text/css" href="../../style.css">
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

<body bgcolor="#000000" text="#FFFFFF">

<table width="744" border="0" cellspacing="0" cellpadding="0" align="center"
height="89">
<tr>
<td width="744" background="../../jh/tiao20.gif" height="83">
<table border="0" height="24" width="90%" cellspacing="0" cellpadding="0"
align="center">
<tbody>
<tr>
<td height="38" width="100%">
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
<td width="475" height="26" align="center"><font size="3"
color="#E18C59"><b>�L�ȴJ��--&nbsp;��ñ</b></font></td>
<td width="104">
<div align="center">

</div>
</td>
</tr>
</table>
</td>
</tr>
</tbody>
</table>
</td>
</tr>
</table>
<table width="738" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td width="17" background="../../jh/tiao10.gif">&nbsp;</td>
<td width="47" valign="top">
<div align="center">
<img src="../../jh/17.gif" width="47" height="251">
</div>
</td>
<td valign="top" width="639">
<div align="center">
<div align="center">
<center>
<div align="center">
<table border="0" width="468" cellspacing="1" cellpadding="0"
height="20">
</center>
</table>
</div>
</div>
<div align="center">
<center>
<table border="0" width="635" cellspacing="1" cellpadding="0">
<tr>
<td class="p1" width="630">��-�z��e����m-�L�ȴJ��--&nbsp;�ڪ���ñ&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;<a href="javascript:history.back(1)"><font color="#FFCC00">��^</font></a></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<form method="POST" action="post.asp">
<table border="0" width="630" cellspacing="1" cellpadding="0">
<tr>
<td class="p3" width="10"></td>
<td class="p3" width="114">�����W�١G</td>
<td class="p3" width="281"><input type="text" name="name"
size="32"></td>
<td class="p3" width="100">��������G</td>
<td class="p3" width="265"><input type="text" name="emptytype"
size="32"></td>
</tr>
<tr>
<td class="p3" width="10"></td>
<td class="p3" width="114">��ܤ����G</td>
<td class="p3" width="281"><select size="1" name="type">
<option value="�T�֥�~">�T�֥�~</option>
<option value="�ӷ~�g��">�ӷ~�g��</option>
<option value="�q������">�q������</option>
<option value="�������d">�������d</option>
<option value="������N">������N</option>
<option value="�Ш|�N�~">�Ш|�N�~</option>
<option value="��ǧ޳N">��ǧ޳N</option>
<option value="��|����">��|����</option>
<option value="���|���">���|���</option>
<option value="�ȹC��q">�ȹC��q</option>
<option value="�F�k�x��">�F�k�x��</option>
<option value="�ͬ��A��">�ͬ��A��</option>
<option value="���|���">���|���</option>
<option value="�ѦҸ��">�ѦҸ��</option>
<option value="�s�D�C��">�s�D�C��</option>
<option value="�C���Ѧa">�C���Ѧa</option>
<option value="���ֺ���">���ֺ���</option>
<option value="�j������">�j������</option>
<option value="�ϮѺ���">�ϮѺ���</option>
<option value="�n��U��">�n��U��</option>
<option value="�q�l�l�c">�q�l�l�c</option>
<option value="�K�O�귽">�K�O�귽</option>
<option value="�Ϯw����">�Ϯw����</option>
<option value="�w��Ѧa">�w��Ѧa</option>
<option value="��ѥ��">��ѥ��</option>
<option value="�����s�{">�����s�{</option>
<option value="���Ͻ׾�">���Ͻ׾�</option>
<option value="���W�о�">���W�о�</option>
<option value="�q�l�P�d">�q�l�P�d</option>
</select></td>
<td class="p3" width="100">�����a�}�G</td>
<td class="p3" width="265">http://<input type="text" name="url"
size="27"></td>
</tr>
<tr>
<td class="p3" width="10"></td>
<td class="p3" width="114">�O�_�}��G</td>
<td class="p3" width="281"><input type="radio" value="�O"
checked name="open">�O<input type="radio" name="open"
value="�_">�_</td>
<td class="p3" width="100"></td>
<td class="p3" width="265"><input type="submit" value="����"
name="B1" style="font-family: Tahoma; font-size: 12px"><input
type="reset" value="�������g" name="B2"
style="font-family: Tahoma; font-size: 12px"></td>
</tr>
</table>
</form>
</center>
</div>
<div align="center">
<center>
<table border="0" width="630" cellspacing="1" cellpadding="0">
<tr>
<td class="p1" width="630">&nbsp;</td>
</tr>
</table>
</center>
</div>
</div>
</td>
<td width="25" background="../../jh/tiao10.gif"> </td>
</tr>
</table>
<table width="731" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td>
<div align="right">
<img src="../../jh/tiao19.gif" width="728" height="31">
</div>
</td>
</tr>
</table>
<div align="center">
</div>

</body>

</html>
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

<%
rs.open "select * from bookmark_count order by count desc",conn,1,1
if rs.bof then
rs.close
%>
<script language="Vbscript">
msgbox"�S���}�񪺦��ç��I",0,"����"
history.back
</script>
<%end if%>
<%
pagecount=20

dqpage=cint(Request.QueryString("page"))
if dqpage=0 then
dqpage=1
end if

temp1=rs.recordcount/pagecount
if temp1-int(rs.recordcount/pagecount)<>0 then
totalpage=int(temp1+1)
else
totalpage=temp1
end if
if dqpage>totalpage then
dqpage=1
end if

if dqpage<>1 then
temp2=(dqpage-1)*pagecount
rs.move temp2

end if
%>
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
<td width="475" height="26">
<div align="center">
<font size="2" color="#E18C59"><b>�ӫ���ñ</b></font>
</div>
</td>
<td width="104">
<div align="center">
<a href="../../jh.asp" target="_top"><font color="#E18C59">��^�ӫ�����</font></a>
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
<td class="p1" width="630">��-�z��e����m--�L�ȴJ��---��ñ����&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;<a href="javascript:history.back(1)"><font color="#FFCC00">��^</font></a></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="630" cellspacing="1" cellpadding="0">
<tr>
<td class="p2" width="10"></td>
<td class="p2" width="40">
<%if dqpage>1 then%>
<form method="get" action="Marklist222.asp">
<input type="hidden" name="page" value="1"><input
type="submit" value="����"
style="font-family: Tahoma; font-size: 12px">
</form>

<%else%>
<input type="submit" value="����"
style="font-family: Tahoma; font-size: 12px">


<%end if%>
</td>
<td class="p2" width="40">
<%if dqpage>1 then%>

<form method="get" action="Marklist222.asp">
<input type="hidden" name="page" value="<%=dqpage-1%>"><input
type="submit" value="�W��"
style="font-family: Tahoma; font-size: 12px">
</form>
<%else%>
<input type="submit" value="�W��"
style="font-family: Tahoma; font-size: 12px">
<%end if%> </td>
<td class="p2" width="40">
<%if dqpage<totalpage then%>
<form method="get" action="Marklist222.asp">
<input type="hidden" name="page" value="<%=dqpage+1%>"><input
type="submit" value="�U��"
style="font-family: Tahoma; font-size: 12px">
</form>

<%else%>
<input type="submit" value="�U��"
style="font-family: Tahoma; font-size: 12px">


<%end if%>
</td>
<td class="p2" width="40">
<%if dqpage<totalpage then%>
<form method="get" action="Marklist222.asp">
<input type="hidden" name="page" value="<%=totalpage%>"><input
type="submit" value="����"
style="font-family: Tahoma; font-size: 12px">
</form>

<%else%>
<input type="submit" value="����"
style="font-family: Tahoma; font-size: 12px">
<%end if%> </td>
<td class="p2" width="446">&nbsp;&nbsp;&nbsp; <a
href="useropen.asp"><font color="#FFCC00">��ñ�`�Ʀ�</font></a></td>
<td class="p2" width="80">��e���G<%=dqpage%>
</td>
<td class="p2" width="80">�`���G<%=totalpage%>
</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="630" cellspacing="1" cellpadding="0">
<tr>
<td class="p2" width="10"></td>
<td class="p2" width="100" align="center">�Ʀ�W��</td>

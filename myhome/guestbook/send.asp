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
<title>�F�C����-�d���ө�</title>
<link rel="stylesheet" type="text/css" href="../../style.css">
<script language="JavaScript">
<!--
function view(newsfile){var gt=unescape('%3e');var popup=null;var over="Launch Pop-up Navigator";popup=window.open('','popupnav','width=510,height=470,resizable=1,scrollbars=yes');if(popup!=null){if(popup.opener==null){popup.opener=self;}popup.location.href=newsfile;}}//-->
</script>

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
<font size="2" color="#E18C59"><b> ���B�ͯd��</b></font>
</div>
</td>
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
<table border="1" width="695" cellspacing="1" cellpadding="0"
bordercolor="#E18C59">
<tr>
<td class="p1" width="476">��-�z��e����m---�L�ȴJ��---�d���O----���B�ͯd��</td>
<td width="10">&nbsp;</td>
<td class="p1" width="290">��-��e�ɶ��G<%=date%><%=time%></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="690" cellspacing="0" cellpadding="0">
<tr>
<td width="100%"></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="1" width="690" cellspacing="1" cellpadding="0"
bordercolor="#E18C59">
<tr>
<td class="p1" width="400">��-���ڪ��̷s�d��</td>
<td class="p1" width="78">��-<a href="javascript:history.back(1)"><font
color="#FFCC00">��^</font></a></td>
<td width="10">&nbsp;</td>
<td class="p1" width="290">��-�ڷQ���ڪ��B�ͯd��</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="690" cellspacing="0" cellpadding="0">
<tr>
<td class="p3" width="474" valign="top">
<table border="1" width="431" cellspacing="1" cellpadding="0"
bordercolor="#E18C59">
<tr>
<td class="p2" width="203">��-�o�����e</td>
<td class="p2" width="67">��-�o���H</td>
<td class="p2" width="100">��-�o���ɶ�</td>
<td class="p2" width="47">���A</td>
<td class="p2" width="44">���v</td>
</tr>
<%
rs.open "select * from guestmessage where receive='"&info(0)&"' and messagetype='�@�����' order by state desc",conn,1,1
count=0
do while not rs.eof and count<8
%>
<% message=replace(rs("message"),"<br>","")
if len(message)>14 then
message=left(message,14)&" ..."
end if%>
<tr>
<td class="p3" width="203">��&nbsp;<a
href="read.asp?id=<%=rs("id")%>"><%=message%></a></td>
<td class="p3" width="67"><%=rs("send")%></td>
<td class="p3" width="100"><%=left(rs("date"),10)%></td>
<td class="p3" width="47"><%
if rs("state")="1" then
Response.Write( "<b>��Ū</b>")
else
Response.Write( "�wŪ")
end if
%></td>
<td class="p3" width="44"><a
href="history.asp?send=<%=rs("send")%>"><font
color="#FFCC00">�I��</font></a></td>
</tr>
<%
count=count+1
rs.movenext
loop
if count<8 then
for i=1 to 8-count
%>

<tr>
<td class="p3" width="203">&nbsp;</td>
<td class="p3" width="67"> </td>
<td class="p3" width="100"> </td>
<td class="p3" width="47"> </td>
<td class="p3" width="44"> </td>
</tr>
<%
next
end if
rs.close
%>

<tr>
<td class="p2" width="461" colspan="5" align="right"><a
href="disp_guest.asp"><font color="#FFCC00">��h�d��...</font></a></td>
</tr>
</table>
<table border="1" width="429" cellspacing="1" cellpadding="0"
bordercolor="#E18C59">
<tr>
<td class="p3" width="429">&nbsp;</td>
</tr>
<tr>
<td class="p3" width="429"> </td>
</tr>
<tr>
<td class="p3" width="429">&nbsp;&nbsp;</td>
</tr>
<tr>
<td class="p3" width="429">&nbsp;&nbsp;&nbsp;</td>
</tr>
</table>
</td>
<td width="10"></td>
<td width="290" valign="top">
<%
send=request.querystring("name")
%>
<table border="1" width="169" cellspacing="1" cellpadding="0"
height="166" bordercolor="#E18C59">
<tr>
<form method="POST" action="post.asp">
<td class="p3" width="246" colspan="2" height="15">��-�o�e�d�����G<a
href="../friend/userinfo.asp?id=<%=send%>"><%=send%></a></td>
</tr>
<input type="hidden" name="name" value="<%=send%>">
<tr>
<td class="p3" width="246" colspan="2" height="16">��-�d�����e�G</td>
</tr>
<tr>
<td width="246" colspan="2" height="100"><textarea rows="4"
name="message" cols="37"></textarea></td>
</tr>
<tr>
<td class="p3" width="144" height="27">&nbsp;<input
type="checkbox" name="receipt" value="ON">
�O�_�ݭn�^���H</td>
<td class="p3" width="106" height="27"><input type="submit"
value="����" name="B1"><input type="reset"
value="�������g" name="B2"></td>
</tr>
</form>
</table>
<%
rs.open "select * from usertype where user='"&info(0)&"' and type='��' order by id desc"
friendcount=rs.recordcount
%>
<table border="1" width="142" align="left" cellspacing="1"
cellpadding="0" bordercolor="#E18C59">
<tr>
<td class="p2" width="100%">�ڪ��B��</td>
</tr>
<%
count=0
do while not rs.eof and count<8
%>
<tr>
<td class="p3" width="100%"><a
href="send.asp?name=<%=rs("name")%>"><font color="#FFCC00"><%=rs("name")%></font></a></td>
</tr>
<%
count=count+1
rs.movenext
loop
if count<8 then
for i=1 to 8-count
%>
<tr>
<td class="p3" width="100%">&nbsp;</td>
</tr>
<%
next
end if
rs.close
%>
<tr>
<td class="p2" width="100%"><%if friendcount>0 then%>
<a href="../friend/hylb.asp"><font color="#FFCC00">�޲z�ڪ��B��</font></a><font
color="#FFCC00"> </font>
<%end if%>
</td>
</tr>
</table>
<%
rs.open "select * from usertype where user='"&info(0)&"' and type='��' order by id desc"
blackcount=rs.recordcount
%>

<table border="1" width="142" cellspacing="1" cellpadding="0"
bordercolor="#E18C59">
<tr>
<td class="p2" width="100%">�¦W��</td>
</tr>
<%
count=0
do while not rs.eof and count<8
%>

<tr>
<td class="p3" width="100%"><a
href="send.asp?name=<%=rs("name")%>"><%=rs("name")%></a></td>
</tr>
<%
count=count+1
rs.movenext
loop
if count<8 then
for i=1 to 8-count
%>

<tr>
<td class="p3" width="100%">&nbsp;</td>
</tr>
<%
next
end if
rs.close
%>
<tr>
<td class="p2" width="100%"><%if blackcount>0 then%>
<a href="friend/black.asp"><font color="#FFCC00">�޲z�ڪ��¦W��</font></a><font
color="#FFCC00"> </font>
<%end if%> </td>
</tr>
</table>
</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="690" cellspacing="0" cellpadding="0">
<tr>
<td class="p1" width="100%">&nbsp;</td>
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
<%conn.close%>

</html>

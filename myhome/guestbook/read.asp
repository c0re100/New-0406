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
<%
rs.open "select * from guestmessage where id="&request.querystring("id"),conn,1,3
if rs.bof then
%>

<script language="Vbscript">
msgbox "�H�����s�b",0,"����"
history.back
</script>
<%
rs.close
conn.close
Response.End
end if

%>

<html>

<head>
<title>�F�C����-�\Ū�d��</title>
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
<font size="2" color="#E18C59"><b>�\Ū�d��</b></font>
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
<td class="p1" width="476">��-�z��e����m- --�L�ȴJ��---�d���O---�\Ū�d��</td>
<td width="10"> </td>
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
<td class="p1" width="76">��-<a href="javascript:history.back(1)"><font
color="#FFCC00">��^</font></a></td>
<td width="10"> </td>
<td class="p1" width="290">��-�ڷQ���ڪ��B�ͯd��</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="690" cellspacing="0" cellpadding="0">
<tr>
<td width="474" valign="top">
<table border="1" width="443" cellspacing="1" cellpadding="0"
bordercolor="#E18C59">
<tr>
<td class="p2" width="100">��-�o�e�H�G</td>
<td class="p3" width="333"><%=rs("send")%></td>
</tr>
<tr>
<td class="p2" width="100">��-�o�e�ɶ��G</td>
<td class="p3" width="333"><%=rs("date")%> </td>
</tr>
<tr>
<td class="p2" width="100">��-���������G</td>
<td class="p3" width="333"><%=rs("messagetype")%> </td>
</tr>
<tr>
<td class="p2" width="100"> </td>
<td class="p3" width="333"> </td>
</tr>
<tr>
<td class="p2" width="100">��-�ԲӤ��e�G</td>
<td class="p3" width="333"><%=rs("message")%> </td>
</tr>
</table>
<%
send=rs("send")
tempmessage=rs("message")
rs.update "state","0"

if rs("receipt")="�O" then
rs.update "receipt","�_"
rs.close
rs.open "select * from guestmessage where id="&request.querystring("id"),conn,1,3
rs.addnew
rs.movelast
rs.update "send","Ū���T�{"
rs.update "receive",send
rs.update "date",ccdate(date)&" "&cctime(time)
rs.update "state","1"
rs.update "receipt","�_"
rs.update "messagetype","Ū���T�{"
message="�z������ "&info(0)&" ���d���w�g�QŪ���A�z���d���O�o�˪��G<br>"
message=message&"-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-<br>"
message=message&tempmessage
conn.execute "update guestmessage set message='"&message&"' where id="&rs("id")
end if
rs.close

if trim(send)<>"Ū���T�{" and trim(send)<>"�t�ή���" and trim(send)<>info(0) then
%>

<table border="1" width="286" cellspacing="1" cellpadding="0"
bordercolor="#E18C59">
<tr>
<td class="p1" width="101" bgcolor="#FFCB31">
<form method="POST" action="post.asp">
<p><font color="#000000">���^�_�d�����G</font></p>
</td>
<td class="p1" width="132" bgcolor="#FFCB31"><%=send%>

</td>
<input type="hidden" name="name" value="<%=send%>">
<td class="p1" width="237" bgcolor="#FFCB31"> </td>
</tr>
<tr>
<td class="p3" width="472" colspan="3">��-�d�����e�G<br>
<textarea rows="6" name="message" cols="70"
style="font-family: Tahoma; font-size: 12px"></textarea></td>
</tr>
<tr>
<td class="p3" width="236" colspan="2">
<p><input type="checkbox" name="receipt" value="ON"> <font
color="#FFCC00">�O�_�ݭn�^���H</font></p>
</td>
<td class="p3" width="240"><input type="submit"
value="�o�e�d��" name="B1"></td>
</form>
</tr>
</table>
<table border="0" width="443" cellspacing="1" cellpadding="0">
<tr>
<td class="p3" width="441">&nbsp;</td>
</tr>
</table>
<table border="1" width="443" cellspacing="1" cellpadding="0"
bordercolor="#E18C59" height="61">
<tr>
<td class="p3" width="443" height="18">��-<a
href="../friend/addusertype.asp?type=��&amp;name=<%=send%>"><font
color="#FFCC00">�[���n��</font></a></td>
</tr>
<tr>
<td class="p3" width="443" height="13">��-<a
href="../friend/addusertype.asp?type=��&amp;name=<%=send%>"><font
color="#FFCB31">�[�J�¦W��</font></a></td>
</tr>
<tr>
<td class="p3" width="443" height="18">&nbsp;</td>
</tr>
</table>
<%end if%>
</td>
<td width="10"></td>
<td width="290" valign="top">
<table border="1" width="1" cellspacing="1" cellpadding="0"
bordercolor="#E18C59">
<tr>
<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript"><!--
function FrontPage_Form2_Validator(theForm)
{

if (theForm.name.value == "")
{
alert("�Цb name �줤��J�ȡC");
theForm.name.focus();
return (false);
}

if (theForm.name.value.length > 50)
{
alert("�b name �줤�A�г̦h��J 50 �Ӧr�šC");
theForm.name.focus();
return (false);
}
return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="post.asp" onsubmit="return FrontPage_Form2_Validator(this)" name="FrontPage_Form2">
<td class="p3" width="248" colspan="2">��-�ڪB�ͪ��W�r�G<!--webbot
bot="Validation" b-value-required="TRUE"
i-maximum-length="50" --><input type="text" name="name"
size="17" maxlength="50"></td>
</tr>
<tr>
<td class="p3" width="248" colspan="2">��-�d�����e�G</td>
</tr>
<tr>
<td width="248" colspan="2"><textarea rows="4"
name="message" cols="37"></textarea></td>
</tr>
<tr>
<td class="p3" width="144"><input type="checkbox"
name="receipt" value="ON"> �O�_�ݭn�^���H</td>
<td class="p3" width="108"><input type="submit"
value="����" name="B1"><input type="reset"
value="�������g" name="B2"></td>
</tr>
</form>
</table>

<%
rs.open "select * from usertype where user='"&info(0)&"' and type='��' order by id desc",conn,1,1
friendcount=rs.recordcount
%>
<table border="1" width="142" align="left" cellspacing="1"
cellpadding="0" height="72" bordercolor="#E18C59">
<tr>
<td class="p2" width="100%" height="16">��-�ڪ��B��</td>
</tr>
<%
count=0
do while not rs.eof and count<8
%>
<tr>
<td class="p3" width="100%" height="16"><a
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
<td class="p3" width="100%" height="16">&nbsp;</td>
</tr>
<%
next
end if
rs.close
%>
<tr>
<td class="p2" width="100%" height="16"><%if friendcount>0 then%>
<a href="../friend/hylb.asp"><font color="#FFCC00">�޲z�ڪ��B��</font></a><font
color="#FFCC00"><a href="more_friend.asp">...</a></font><%end if%></td>
</tr>
</table>
<%
rs.open "select * from usertype where user='"&info(0)&"' and type='��' order by id desc"
blackcount=rs.recordcount
%>

<table border="1" width="142" cellspacing="1" cellpadding="0"
bordercolor="#E18C59">
<tr>
<td class="p2" width="100%">��-�¦W��</td>
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
<a href="../friend/black.asp"><font color="#FFCC00">�޲z�ڪ��¦W��</font></a><a
href="more_black.asp"><font color="#FFCC00">..</font>..</a>
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
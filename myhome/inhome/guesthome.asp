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
<!--#INCLUDE file="check.asp"-->
<%
cname=check(request("name"),"�W�r")
%>
<%
hostname=request.querystring("name")
guestname=session("user")
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�F�C����V�X��-���U</title>
<link rel="StyleSheet" href="new.css" title="Contemporary">
</head>

<body topmargin="0" leftmargin="0" bgcolor="#000000" text="#E18C59">

<%
rs.open "select * from userinfo where user='"&hostname&"' and homeopen='1'",conn,1,1

if rs.bof then
%>
<script language="Vbscript">
msgbox"�o���Ъ��򳬵ۡA���G�@���S���H��L",0,"����"
history.back
</script>
<%
rs.close
conn.close
Response.End
end if
%>

<div align="center">
<center>
<table border="0" width="776" cellspacing="0" cellpadding="0">
<tr>
<td width="100%">&nbsp;</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="776" cellspacing="1" cellpadding="0">
<tr>
<td class="p1" width="336"><font size="2">��-�z��e����m-<a
href="city.asp"><b><font color="#FFFF00">�ӫ�����</font></b></a>-<%=hostname%>���a</font></td>
<td class="p1" width="440"><font size="2">��-��e�ɶ��G<%=date%><%=time%></font></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="776" cellspacing="1" cellpadding="0">
<tr>
<td class="p2" width="496" colspan="4" height="21"><font size="2">�A�w�g�b<%=hostname%>�a�̤F�A�D�H�{�b�n���S���b�a</font></td>
</tr>
<tr>
<td class="p2" width="102" height="15"><font size="2">&nbsp;�D�H���H���G</font></td>
<td class="p2" width="64" height="15"></td>
<td class="p2" width="68" height="15"><font size="2">&nbsp;�Ӥ��G</font></td>
<td class="p2" width="248" height="18"><font size="2">&nbsp;
�D�H���X�̪��ܡG</font></td>
<td class="p2" width="262" height="18" colspan="3"><font size="2">&nbsp;�D�H�}�񪺰ϰ�</font></td>
</tr>
<tr>
<td class="p3" width="102" height="16"><font size="2">&nbsp;�m�W�G</font></td>
<td class="p3" width="64" height="16"><font size="2"><%=rs("user")%></font></td>
<td class="p3" width="68" height="55" rowspan="4" align="center"><font
size="2"><img
src="../register/skin/<%if rs("skin")<>"" then
skin=rs("skin")
else
skin="01.gif"
end if
response.write skin%>"
alt="<%=rs("user")%>"></font></td>
<td class="p3" width="250" height="151" rowspan="8" valign="top"><font
size="2"><%=rs("guestmemo")%></font></td>
<td class="p3" width="261" height="139" valign="top" colspan="3"
rowspan="8">
<%
rs.close
rs.open "select * from openlist where user='"&hostname&"' and open='�O'",conn,1,1
%>
<table border="0" width="100%" cellspacing="1" cellpadding="0">
<tr>
<td class="p3" width="37%"><a
href="../guestbook/send.asp?name=<%=hostname%>"><font size="2"
color="#FFFF00">���D�H�d��</font></a></td>
<td class="p3" width="63%"></td>
</tr>
<%
do while not rs.eof
%>
<tr>
<td class="p3" width="37%"><a
href="../homejump.asp?fun=<%=rs("function")%>&amp;hostname=<%=hostname%>"><font
size="2"><%=rs("function")%></font></a></td>
<td class="p3" width="63%"></td>
</tr>
<%
rs.movenext
loop%>
</table>
</td>
</tr>
<tr>
<td class="p3" width="102" height="4"><font size="2">&nbsp;�~�֡G</font></td>
<%
rs.close
rs.open "select * from userinfo where user='"&hostname&"'",conn,1,3

%>
</tr>
<tr>
<td class="p3" width="102" height="19"><font size="2">&nbsp;�ʧO�G</font></td>
<td class="p3" width="64" height="19"><font size="2"><%=rs("sex")%> </font></td>
</tr>
<tr>
<td class="p3" width="102" height="5"><font size="2">&nbsp;�ӳX�H���G</font></td>
<%
guestcount=rs("guestcount")+1
rs.update "guestcount",guestcount
%>
<td class="p3" width="132" height="5" colspan="2"><font size="2"><%=guestcount%> </font></td>
</tr>
</table>
</center>
</div>
<%rs.close%>
<div align="center">
<center>
<table border="0" width="776" cellspacing="1" cellpadding="0">
<tr>
<td class="p1" width="776"><font size="2">&nbsp;</font></td>
</tr>
</table>
</center>
</div>

</body>
<%conn.close%>

</html>
<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
name=info(0)
rs.open "SELECT �����[,���s�[,��O�[,���O�[,�m�W,�Ȩ�,allvalue,grade,����,���s,���O,����,��O FROM �Τ� WHERE id="& info(9),conn
gjj=rs("�����[")
fyj=rs("���s�[")
tlj=rs("��O�[")
nlj=rs("���O�[")
%>
<html>
<head>
<title>�ڪ��]��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="style.css">
</head>
<body bgcolor="#FFFFFF" background="back.gif">
<div align="center">
<table width="509" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="73" rowspan="3" valign="top">
<div align="center">
<input onClick="javascript:window.document.location.href='index.asp'" title=�˳Ƥ@�� type=button value="�˳Ƥ@��" name="button7">
<br>
          <input onClick="javascript:window.document.location.href='head.asp'" title=�Z���˳� type=button value="�Z���˳�" name="button">
<br>
          <br>
          <input onClick="javascript:window.document.location.href='eat.asp'" title=����Ī� type=button value="����Ī�" name="button62">
</div>
</td>
<td valign="top" width="436">
<div align="center"><img src="dao.gif" width="40" height="15">�ڪ��˳Ƥ@��<img src="dao1.gif" width="40" height="15">
<font color="#CC0000" face="����"><a href="javascript:this.location.reload()">��s</a></font>
</div>
</td>
</tr>
<tr>
<td valign="top" width="436" height="27">
<div align="center">�Τ�<font color="#FF0000"><%=rs("�m�W")%></font>���A�A���⬰�W���I<font color="#FF0000"><font color="#000000">�Ȥl�G<%=rs("�Ȩ�")%>
�g��G<%=rs("allvalue")%></font></font><font color="#FF0000"><font color="#000000">
���šG<%=rs("grade")%>��</font></font><br>
�����O�G<%=rs("����")%>
���s�O�G<%=rs("���s")%>
���O�G<%=rs("���O")%><font color="#FF0000">/<%=rs("����")*64+2000+nlj%></font>
��O�G<%=rs("��O")%><font color="#FF0000">/<%=rs("����")*240+5260+tlj%>
</font></div>
</td>
</tr>
<tr>
<td width="436" valign="top" height="172">
<table width="78%" border="1" cellpadding="2" cellspacing="0" bordercolor="#000000" align="center">
<tr>
<td> �Y���G
<%
rs.close
rs.open "SELECT ���~�W,���s FROM ���~ WHERE �֦���="& info(9) & " and �˳�=true and ����='�Y��'",conn
if rs.eof or rs.bof then
%>
���˳�
<%else%>
<%=rs("���~�W")%> ���s�G<%=rs("���s")%>
<%end if%>
</td>
</tr>
<tr>
<td> �˹��G
<%
rs.close
rs.open "SELECT ���~�W,���s FROM ���~ WHERE �֦���=" & info(9) & " and �˳�=true and ����='�˹�'",conn
if rs.eof or rs.bof then
%>
���˳�
<%else%>
<%=rs("���~�W")%> ���s�G<%=rs("���s")%>
<%end if%>
</td>
</tr>
<tr>
<td> ���ҡG
<%
rs.close
rs.open "SELECT ���~�W,���s FROM ���~ WHERE �֦���=" & info(9) & " and �˳�=true and ����='����'",conn
if rs.eof or rs.bof then
%>
���˳�
<%else%>
<%=rs("���~�W")%> ���s�G<%=rs("���s")%>
<%end if%>
</td>
</tr>
<tr>
<td height="2"> ����G
<%
rs.close
rs.open "SELECT ���~�W,���s FROM ���~ WHERE �֦���="&info(9)& " and �˳�=true and ����='����'",conn
if rs.eof or rs.bof then
%>
���˳�
<%else%>
<%=rs("���~�W")%> ���s�G<%=rs("���s")%>
<%end if%>
</td>
</tr>
<tr>
<td> �k��G
<%
rs.close
rs.open "SELECT ���~�W,���� FROM ���~ WHERE �֦���=" & info(9) & " and �˳�=true and ����='�k��'",conn
if rs.eof or rs.bof then
%>
���˳�
<%else%>
<%=rs("���~�W")%> �����G<%=rs("����")%>
<%end if%>
</td>
</tr>
<tr>
<td> ���}�G
<%
rs.close
rs.open "SELECT ���~�W,���s FROM ���~ WHERE �֦���=" & info(9) & " and �˳�=true and ����='���}'",conn
if rs.eof or rs.bof then
%>
���˳�
<%else%>
<%=rs("���~�W")%> ���s�G<%=rs("���s")%>
<%end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</td>
</tr>
</table>
</td>
</tr>
</table>
<table width="475" border="0" cellspacing="0" cellpadding="0">
<tr>
<td valign="top" align="center">
<input onClick="JavaScript:window.close()" title=�������f type=button value="�������f" name="button2">
</td>
</tr>
</table>
</div>
</body>
</html>

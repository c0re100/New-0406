<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �|���O from �Τ� where id="&info(9),conn
hyfei=rs("�|���O")
rs.close
rs.open "select ���~�W,����,�Ȩ�,ID from ���~�R where ����='�d��' and �Ȩ�>=1 and �Ȩ�<5000 order by �Ȩ�"
%>
<html>
<head>
<title>�|���d��</title>
<link rel="stylesheet" href="../chat/ccss2.css">
</head>
<body bgcolor=#FFFFFF background="../ff.jpg" text="#000000">
<BR>
<div align="center">
<font color=blue><%=info(0)%></font>�{�b���|���O:<font color=red><b><%=hyfei%></b></font>����
</div>
<table border=1 align=center width=620 cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#3399FF">
<tr>
<td width="10%" height="11"><div align="center">�d���W��</div></td>
<td width="69%" height="11"><div align="center">�\ ��</div></td>
<td width="11%" height="11"><div align="center">�|�O</div></td>
<td height="11" width="10%"><div align="center">�� �@</div></td>
</tr>
<%
do while not rs.bof and not rs.eof
%>
<tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';">
<td><div align="center"><%=rs("���~�W")%> </div></td>
<td><%=rs("����")%></td>
<td><div align="center"><%=rs("�Ȩ�")%>��</div></td>
<td align="center"><a href="buycard.asp?id=<%=rs("id")%>">�ʶR</a></td>
</tr>
<%
rs.movenext
loop
%>
</table><br><br>
<table border=1 align=center width=620 cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFFF">
<%rs.close
rs.open "select ���~�W,����,����,ID from ���~�R where ����='�d��' and ����>=1 order by �Ȩ�"
do while not rs.bof and not rs.eof
%>
<tr  bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';">
<td width='12%'><div align="center"><%=rs("���~�W")%> </div></td>
<td width='65%'><%=rs("����")%></td></td>
<td width='16%'><div align="center"><%=rs("����")%>�Ӫ���</div></td>
<td align="center" width="10%"><a href="buycard.asp?id=<%=rs("id")%>">�ʶR</a></td>
</tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
</body>
</html>
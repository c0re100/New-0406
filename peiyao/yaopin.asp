<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
name=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT ���~�W,�ƶq,�\�� FROM �׽m���~ WHERE �֦���=" & info(9) & " and �ƶq>0  order by �W�[�ƭ�",conn
%>
<html>
<head>
<title>�t�m�ħ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="style.css">
</head>
<body bgcolor="#660000" text="#FFFFFF" leftmargin="0" topmargin="0" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<div align="left">
  <div align="center">�{�w�t�ħ��@�� <font face="����"><a href="javascript:this.location.reload()">��s</a></font></div>
<div align="center"> <br>
<table border="1" align="center" width="141" cellpadding="1" cellspacing="0" height="31">
<tr align="center">
<td nowrap width="64">
<div align="center"><font color="#FFFFFF" size="-1">���~�W</font></div>
</td>
<td nowrap width="33">
<div align="center"><font color="#FFFFFF" size="-1">�ƶq </font> </div>
<td colspan="2" nowrap width="36">
<div align="center"><font color="#FFFFFF" size="-1">�\��</font></div>
</td>
</tr>
<%
do while not rs.eof
%>
<tr>
<td width="64" height="3">
<div align="center"><font color="#FFFFFF" size="-1"><%=rs("���~�W")%>
</font> </div>
</td>
<td width="33" height="3">
<div align="center"><font color="#FFFFFF" size="-1"><%=rs("�ƶq")%>
</font> </div>
<td colspan="2" height="3" width="36">
<div align="center"><font color="#FFFFFF" size="-1"><%=rs("�\��")%>
</font> </div>
</td>
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
</div>
</div>
</body>
</html>

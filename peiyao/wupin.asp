<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(0)="" then Response.Redirect "../error.asp?id=210"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
Response.Expires=0
name=info(0)
rs.open "SELECT ���~�W,�ƶq,�Ȩ� FROM ���~ WHERE �֦���=" & info(9) & " and �ƶq>0  and (���~�W='�j�T��' or ���~�W='�B��' or ���~�W='�q��' or ���~�W='�j��' or ���~�W='�Ѫ��' or ���~�W='�p�U��') order by �Ȩ� ",conn
%>
<html>
<head>
<title>�ħ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="../style.css">
</head>
<body bgcolor="#660000" text="#FFFFFF" leftmargin="0" topmargin="0">
<div align="left">
  <div align="center"><font color="#FFFFFF">�{���ħ��@�� </font><font face="����"><a href="javascript:this.location.reload()"><font color="#FFFF00">��s</font></a></font></div>
<div align="center"> <br>
<table border="1" align="center" width="141" cellpadding="1" cellspacing="0" height="31">
<tr align="center">
<td nowrap width="64" height="16">
<div align="center"><font color="#FFFFFF" size="-1">���~�W</font></div>
</td>
<td nowrap width="33" height="16">
<div align="center"><font color="#FFFFFF" size="-1">�ƶq </font> </div>
<td colspan="2" nowrap width="36" height="16">
<div align="center"><font color="#FFFFFF" size="-1">����</font></div>
</td>
</tr>
<%
do while not rs.eof
%>
<tr>
<form method=POST action='wupin1.asp?action=<%=rs("���~�W")%>&name=<%=name%>'>
<td width="64" height="3">
<div align="center"><font color="#FFFFFF" size="-1"><%=rs("���~�W")%>
</font> </div>
</td>
<td width="33" height="3">
<div align="center"><font color="#FFFFFF" size="-1"><%=rs("�ƶq")%>
</font> </div>
<td colspan="2" height="3" width="36">
<div align="center"><font color="#FFFFFF" size="-1"><%=rs("�Ȩ�")%>
</font> </div>
</td>
</form>
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

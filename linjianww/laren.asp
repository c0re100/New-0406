<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"
larenseek=Request.Form("larenseek")
%>
<html>
<head>
<title>�F�C����|���d�ߵ{��</title>
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: �s�ө���}
.c {  font-family: �s�ө���; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body text="#000000" vlink="#FF9900" topmargin="0"
leftmargin="0" background="0064.jpg">
<p align="center"> <font color="#CC0000" face="����"><a href="javascript:this.location.reload()">��s</a></font>
<br>
�P�§A�B�ͳo�ǤH�O�A�Ԩ�ڭ̦��򪺡I<br>

<table border="0" width="500" cellspacing="0" cellpadding="0"
background="../jhmp/bg.gif" align="center">
  <tr align="center">
    <td background="../jhmp/top1.gif" width="500" height="26"><font
color="#FF6600"><b><font size="+1" color="#FFFF00">����ԤH�O��</font></b></font></td>
</tr>
<tr align="center">
<td>
      <table width="470" border="1" cellpadding="0" cellspacing="0" height="13">
        <tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="SELECT * FROM �Τ� where ����>1 and ���ФH='"& larenseek &"' order by lasttime"
Set Rs=conn.Execute(sql)
%>
<td width="28" height="10">
<div align="center"><font color="#FFFFFF">ID</font></div>
</td>
<td width="47" height="10">
<div align="center"><font color="#FFFFFF">�m�W</font></div>
</td>
<td width="25" height="10">
<div align="center"><font color="#FFFFFF">�ʧO</font></div>
</td>
<td width="63" height="10">
<div align="center"><font color="#FFFFFF">����</font></div>
</td>
<td width="86" height="10">
<div align="center"><font color="#FFFFFF">����</font></div>
</td>
<td width="75" height="10">
<div align="center"><font color="#FFFFFF">�̫�n���ɶ�</font></div>
</td>
<td width="51" height="10">
<div align="center"><font color="#FFFFFF">���򵥯�</font></div>
</td>
<td width="35" height="10">
<div align="center"><font color="#FFFFFF">�n��</font></div>
</td>
</tr>
<%
jl=0
do while not rs.bof and not rs.eof
jl=jl+1
%>
<tr>
<td width="28" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("ID")%></font></div>
</td>
<td width="47" height="30">
<div align="center"><a href="showuser.asp?username=<%=rs("�m�W")%>"><font color="#FF9900"><%=rs("�m�W")%></font></a></div>
</td>
<td width="25" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("�ʧO")%></font></div>
</td>
<td width="63" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("����")%></font></div>
</td>
<td width="86" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("����")%></font></div>
</td>
<td width="75" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("lasttime")%></font></div>
</td>
<td width="51" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("����")%></font></div>
</td>
<td width="35" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("times")%></font></div>
</td>
</tr>
<%
rs.movenext
loop
conn.close
%>
</table>
</td>
</tr>
</table>
<div align="center">
  <p><font color="#000000">�ԤH�`��:</font><b><font color="#0000FF"><%=(jl)%></font></b><font color="#000000">�H</font></p>
</div>
</body>
</html> 
<!--#include file="conn.asp"-->
<%
response.buffer=true
formsize=request.totalbytes
formdata=request.binaryread(formsize)
bncrlf=chrB(13) & chrB(10)
divider=leftB(formdata,clng(instrb(formdata,bncrlf))-1)
datastart=instrb(formdata,bncrlf & bncrlf)+4
dataend=instrb(datastart+1,formdata,divider)-datastart
mydata=midb(formdata,datastart,dataend)
conn.execute("update list set photo ='' where name='"&session("Username")&"'")
set rec=server.createobject("ADODB.recordset")
rec.Open "SELECT * FROM list where name='"&session("Username")&"' ",conn,1,3
rec("photo")=""
rec("photo").appendchunk mydata

rec.update

set rec=nothing
%>
<HTML><HEAD><TITLE><%=session("myname")%>_�Ӥ��W��</TITLE>
<META content="text/html; charset=big5" http-equiv=Content-Type>
<style type="text/css">
<!--td {  font-family: �s�ө���; font-size: 9pt}
body {  font-family: �s�ө���; font-size: 9pt}
select {  font-family: �s�ө���; font-size: 9pt}
A {text-decoration: none; font-family: "�s�ө���"; font-size: 9pt; color: #ff6666}
A:hover {text-decoration: underline; color: #666666; font-family: "�s�ө���"; font-size: 9pt} .txt {  font-family: "�s�ө���"; font-size: 10.5pt}
--></style>

</HEAD>
<body bgcolor="#8d8051" text="#ffffff" link="#ffffff" alink="#ffffff" vlink="#ffffff" leftmargin="0" topmargin="0" background="../ljimage/11.jpg">
<div align="center"><a href=upload.asp>�����N�A�n���s�W��</a> | <a href=javascript:location.reload();>�Ш�s���</a> |<a href="main.asp?action=search">��^��ͦC��</a></div>
<div align="center">
<center>
<br>
<table width="46%" border="1" align="center" cellpadding="1" cellspacing="0" bordercolorlight="#CCCC99" bordercolordark="#996633">
<tr bgcolor="#8d8051">
<td colspan="3">
<div align="right"><a href=javascript:window.close();><font color="#ffffff">��������</a></div>
</td>
</tr>
<tr bgcolor="#b2a265">
<td colspan="3">
<div align="center"> <font color="#000000"><%=session("myname")%>���Ӥ��W�Ǧ��\</font></div>
</td>
</tr>
<tr>
<td height="193" width="25%" bgcolor="#8d8051">���j�p:<br><%=formsize%>�r�`</td>
<td height="193" width="75%" bgcolor="#8d8051">
<div align="center">
<img src="showimg.asp?id=<%=session("Username")%>" align="middle"></div>
</td>
</tr>
</table>
</center>
</div><p>
<div align="center"><a href=upload.asp>�����N�A�n���s�W��</a> | <a href=javascript:location.reload();>�Ш�s���</a> |<a href="main.asp?action=search">��^��ͦC��</a></div>
</body>
</html>
<%
'end if
'end if
%>


